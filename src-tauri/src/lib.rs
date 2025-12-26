use base64::{engine::general_purpose, Engine as _};
use image::codecs::jpeg::JpegEncoder;
use image::codecs::webp::WebPEncoder;
use image::{imageops::FilterType, ColorType, DynamicImage, GenericImageView};
use serde::Serialize;
use tauri::Emitter;
use std::collections::HashMap;
use std::io::{self, Read};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

#[derive(Serialize)]
struct CompressionResult {
    data_url: String,
    mime_type: String,
    bytes: usize,
    width: u32,
    height: u32,
}

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
fn compress_receipt_image(data_url: String, max_dim: Option<u32>) -> Result<CompressionResult, String> {
    let (original_mime, original_bytes) = parse_data_url(&data_url)?;
    let decoded = match image::load_from_memory(&original_bytes) {
        Ok(image) => image,
        Err(_) => {
            let data_url = build_data_url(&original_mime, &original_bytes);
            return Ok(CompressionResult {
                data_url,
                mime_type: original_mime,
                bytes: original_bytes.len(),
                width: 0,
                height: 0,
            });
        }
    };
    let (orig_width, orig_height) = decoded.dimensions();

    let resized = resize_image(decoded, max_dim.unwrap_or(2000));
    let (out_width, out_height) = resized.dimensions();

    let mut candidates: Vec<(String, Vec<u8>)> = Vec::new();
    if let Ok(bytes) = encode_webp_lossless(&resized) {
        candidates.push(("image/webp".to_string(), bytes));
    }
    for quality in [85_u8, 75, 65] {
        if let Ok(bytes) = encode_jpeg(&resized, quality) {
            candidates.push(("image/jpeg".to_string(), bytes));
        }
    }

    let mut best_mime = original_mime;
    let mut best_bytes = original_bytes;
    let mut best_width = orig_width;
    let mut best_height = orig_height;

    for (mime_type, bytes) in candidates {
        if bytes.len() < best_bytes.len() {
            best_mime = mime_type;
            best_bytes = bytes;
            best_width = out_width;
            best_height = out_height;
        }
    }

    let data_url = build_data_url(&best_mime, &best_bytes);
    Ok(CompressionResult {
        data_url,
        mime_type: best_mime,
        bytes: best_bytes.len(),
        width: best_width,
        height: best_height,
    })
}

fn parse_data_url(input: &str) -> Result<(String, Vec<u8>), String> {
    let input = input.trim();
    if let Some(comma) = input.find(',') {
        let (meta, data) = input.split_at(comma);
        let meta = meta.trim_start_matches("data:");
        let mime = meta
            .split(';')
            .next()
            .unwrap_or("application/octet-stream")
            .to_string();
        let mut cleaned = data[1..].to_string();
        cleaned.retain(|c| !c.is_whitespace());
        let bytes = general_purpose::STANDARD
            .decode(cleaned)
            .map_err(|e| format!("Base64 decode failed: {e}"))?;
        return Ok((mime, bytes));
    }

    let mut cleaned = input.to_string();
    cleaned.retain(|c| !c.is_whitespace());
    let bytes = general_purpose::STANDARD
        .decode(cleaned)
        .map_err(|e| format!("Base64 decode failed: {e}"))?;
    Ok(("application/octet-stream".to_string(), bytes))
}

fn build_data_url(mime_type: &str, bytes: &[u8]) -> String {
    let encoded = general_purpose::STANDARD.encode(bytes);
    format!("data:{mime_type};base64,{encoded}")
}

fn resize_image(image: DynamicImage, max_dim: u32) -> DynamicImage {
    if max_dim == 0 {
        return image;
    }
    let (width, height) = image.dimensions();
    if width <= max_dim && height <= max_dim {
        return image;
    }
    let scale = (max_dim as f32 / width as f32).min(max_dim as f32 / height as f32);
    let new_width = (width as f32 * scale).round().max(1.0) as u32;
    let new_height = (height as f32 * scale).round().max(1.0) as u32;
    image.resize(new_width, new_height, FilterType::Lanczos3)
}

fn encode_webp_lossless(image: &DynamicImage) -> Result<Vec<u8>, String> {
    let rgba = image.to_rgba8();
    let (width, height) = rgba.dimensions();
    let mut out = Vec::new();
    let encoder = WebPEncoder::new_lossless(&mut out);
    encoder
        .encode(rgba.as_raw(), width, height, ColorType::Rgba8)
        .map_err(|e| e.to_string())?;
    Ok(out)
}

fn encode_jpeg(image: &DynamicImage, quality: u8) -> Result<Vec<u8>, String> {
    let rgb = image.to_rgb8();
    let (width, height) = rgb.dimensions();
    let mut out = Vec::new();
    let mut encoder = JpegEncoder::new_with_quality(&mut out, quality);
    encoder
        .encode(rgb.as_raw(), width, height, ColorType::Rgb8)
        .map_err(|e| e.to_string())?;
    Ok(out)
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
            compress_receipt_image,
            upload_payment_attachment,
            cancel_upload
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
