use serde::Serialize;
use tauri::Emitter;
use std::collections::HashMap;
use std::io::{self, Read};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct UploadProgress {
    upload_id: String,
    loaded: u64,
    total: u64,
    done: bool,
}

#[derive(Default)]
struct UploadState {
    cancel_flags: Mutex<HashMap<String, Arc<AtomicBool>>>,
}

impl UploadState {
    fn register(&self, upload_id: &str) -> Arc<AtomicBool> {
        let mut guard = self.cancel_flags.lock().unwrap();
        let flag = Arc::new(AtomicBool::new(false));
        guard.insert(upload_id.to_string(), flag.clone());
        flag
    }

    fn cancel(&self, upload_id: &str) -> bool {
        let guard = self.cancel_flags.lock().unwrap();
        if let Some(flag) = guard.get(upload_id) {
            flag.store(true, Ordering::SeqCst);
            return true;
        }
        false
    }

    fn remove(&self, upload_id: &str) {
        let mut guard = self.cancel_flags.lock().unwrap();
        guard.remove(upload_id);
    }
}

struct ProgressReader<R: Read> {
    inner: R,
    total: u64,
    sent: u64,
    last_emit: u64,
    emit_every: u64,
    app: tauri::AppHandle,
    upload_id: String,
    cancel_flag: Arc<AtomicBool>,
}

impl<R: Read> ProgressReader<R> {
    fn new(
        inner: R,
        total: u64,
        app: tauri::AppHandle,
        upload_id: String,
        cancel_flag: Arc<AtomicBool>,
    ) -> Self {
        Self {
            inner,
            total,
            sent: 0,
            last_emit: 0,
            emit_every: 64 * 1024,
            app,
            upload_id,
            cancel_flag,
        }
    }

    fn emit(&mut self, done: bool) {
        let payload = UploadProgress {
            upload_id: self.upload_id.clone(),
            loaded: self.sent,
            total: self.total,
            done,
        };
        let _ = self.app.emit("upload-progress", payload);
        self.last_emit = self.sent;
    }
}

impl<R: Read> Read for ProgressReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        if self.cancel_flag.load(Ordering::SeqCst) {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "cancelled"));
        }
        let read = self.inner.read(buf)?;
        if read == 0 {
            self.emit(true);
            return Ok(0);
        }
        self.sent = self.sent.saturating_add(read as u64);
        if self.sent - self.last_emit >= self.emit_every || self.sent >= self.total {
            self.emit(self.sent >= self.total);
        }
        Ok(read)
    }
}

#[tauri::command]
fn upload_payment_attachment(
    app: tauri::AppHandle,
    state: tauri::State<UploadState>,
    url: String,
    payload: serde_json::Value,
    upload_id: String,
) -> Result<serde_json::Value, String> {
    if url.trim().is_empty() {
        return Err("Missing Apps Script URL".to_string());
    }
    let body = serde_json::json!({
        "action": "uploadPaymentAttachment",
        "payload": payload,
    });
    let body_bytes = serde_json::to_vec(&body).map_err(|e| e.to_string())?;
    let total = body_bytes.len() as u64;
    let cancel_flag = state.register(&upload_id);

    let result = (|| {
        let reader = ProgressReader::new(
            io::Cursor::new(body_bytes),
            total,
            app.clone(),
            upload_id.clone(),
            cancel_flag.clone(),
        );
        let response = ureq::post(&url)
            .set("Content-Type", "text/plain")
            .set("Content-Length", &total.to_string())
            .send(reader)
            .map_err(|e| e.to_string())?;
        let text = response.into_string().map_err(|e| e.to_string())?;
        serde_json::from_str(&text).map_err(|e| e.to_string())
    })();

    state.remove(&upload_id);
    result
}

#[tauri::command]
fn cancel_upload(state: tauri::State<UploadState>, upload_id: String) -> bool {
    state.cancel(&upload_id)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_fs::init())
        .manage(UploadState::default())
        .invoke_handler(tauri::generate_handler![
            upload_payment_attachment,
            cancel_upload
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
