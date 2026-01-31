use rusqlite::{params, Connection};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::io::{self, Read, Write};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};
use tauri::{Emitter, Manager};

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct UploadProgress {
    upload_id: String,
    loaded: u64,
    total: u64,
    done: bool,
}

#[derive(Serialize, Clone)]
#[serde(rename_all = "camelCase")]
struct DownloadProgress {
    download_id: String,
    loaded: u64,
    total: u64,
    done: bool,
}

#[derive(Default)]
struct UploadState {
    cancel_flags: Mutex<HashMap<String, Arc<AtomicBool>>>,
}

struct DbState {
    conn: Mutex<Connection>,
}

#[derive(Serialize)]
struct CacheEntry {
    value: serde_json::Value,
    updated_at: i64,
}

#[derive(Deserialize)]
struct CacheSetItem {
    key: String,
    value: serde_json::Value,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SyncJob {
    id: i64,
    action: String,
    payload: serde_json::Value,
    method: String,
    params: serde_json::Value,
    created_at: i64,
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

fn now_ms() -> i64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as i64)
        .unwrap_or(0)
}

fn get_db_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().app_data_dir().map_err(|e| e.to_string())?;
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.join("rent_software.sqlite"))
}

fn init_db(conn: &Connection) -> Result<(), rusqlite::Error> {
    conn.execute_batch(
        "CREATE TABLE IF NOT EXISTS local_cache (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at INTEGER NOT NULL
        );
        CREATE TABLE IF NOT EXISTS sync_queue (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT NOT NULL,
            method TEXT NOT NULL,
            params TEXT NOT NULL,
            payload TEXT NOT NULL,
            created_at INTEGER NOT NULL
        );",
    )?;
    Ok(())
}

fn setup_db(app: &tauri::AppHandle) -> Result<Connection, String> {
    let path = get_db_path(app)?;
    let conn = Connection::open(path).map_err(|e| e.to_string())?;
    conn.execute_batch(
        "PRAGMA journal_mode=WAL;
         PRAGMA synchronous=NORMAL;
         PRAGMA temp_store=MEMORY;
         PRAGMA foreign_keys=ON;",
    )
    .map_err(|e| e.to_string())?;
    init_db(&conn).map_err(|e| e.to_string())?;
    Ok(conn)
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

fn sanitize_segment(value: &str, fallback: &str) -> String {
    let cleaned: String = value
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '-' || ch == '.' {
                ch
            } else {
                '_'
            }
        })
        .collect();
    let trimmed = cleaned.trim_matches('_');
    if trimmed.is_empty() {
        fallback.to_string()
    } else {
        trimmed.to_string()
    }
}

#[tauri::command]
fn cache_get(state: tauri::State<DbState>, key: String) -> Result<Option<CacheEntry>, String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let mut stmt = conn
        .prepare("SELECT value, updated_at FROM local_cache WHERE key = ?1")
        .map_err(|e| e.to_string())?;
    let mut rows = stmt.query(params![key]).map_err(|e| e.to_string())?;
    if let Some(row) = rows.next().map_err(|e| e.to_string())? {
        let value_raw: String = row.get(0).map_err(|e| e.to_string())?;
        let updated_at: i64 = row.get(1).map_err(|e| e.to_string())?;
        let value = serde_json::from_str(&value_raw).unwrap_or(serde_json::Value::Null);
        return Ok(Some(CacheEntry { value, updated_at }));
    }
    Ok(None)
}

#[tauri::command]
fn cache_set(
    state: tauri::State<DbState>,
    key: String,
    value: serde_json::Value,
) -> Result<(), String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let payload = serde_json::to_string(&value).map_err(|e| e.to_string())?;
    let updated_at = now_ms();
    conn.execute(
        "INSERT OR REPLACE INTO local_cache (key, value, updated_at) VALUES (?1, ?2, ?3)",
        params![key, payload, updated_at],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn ensure_temp_dir(
    app: tauri::AppHandle,
    tenant_name: String,
) -> Result<serde_json::Value, String> {
    let base = app.path().temp_dir().map_err(|e| e.to_string())?;
    let tenant = sanitize_segment(&tenant_name, "Tenant");
    let dir = base.join("Tenant_Docs").join(tenant);
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(serde_json::json!({ "ok": true, "path": dir.to_string_lossy().to_string() }))
}

#[tauri::command]
fn cache_set_many(
    state: tauri::State<DbState>,
    entries: Vec<CacheSetItem>,
) -> Result<bool, String> {
    if entries.is_empty() {
        return Ok(true);
    }
    let mut conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let tx = conn.transaction().map_err(|e| e.to_string())?;
    let updated_at = now_ms();
    for entry in entries {
        if entry.key.trim().is_empty() {
            continue;
        }
        let payload = serde_json::to_string(&entry.value).map_err(|e| e.to_string())?;
        tx.execute(
            "INSERT OR REPLACE INTO local_cache (key, value, updated_at) VALUES (?1, ?2, ?3)",
            params![entry.key, payload, updated_at],
        )
        .map_err(|e| e.to_string())?;
    }
    tx.commit().map_err(|e| e.to_string())?;
    Ok(true)
}

#[tauri::command]
fn cache_delete(state: tauri::State<DbState>, key: String) -> Result<(), String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    conn.execute("DELETE FROM local_cache WHERE key = ?1", params![key])
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn cache_delete_prefix(state: tauri::State<DbState>, prefix: String) -> Result<(), String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let pattern = format!("{}%", prefix);
    conn.execute(
        "DELETE FROM local_cache WHERE key LIKE ?1",
        params![pattern],
    )
    .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn queue_add(
    state: tauri::State<DbState>,
    action: String,
    payload: serde_json::Value,
    method: Option<String>,
    params: Option<serde_json::Value>,
) -> Result<i64, String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let payload_json = serde_json::to_string(&payload).map_err(|e| e.to_string())?;
    let method_value = method.unwrap_or_else(|| "POST".to_string());
    let params_value = params.unwrap_or_else(|| serde_json::json!({}));
    let params_json = serde_json::to_string(&params_value).map_err(|e| e.to_string())?;
    let created_at = now_ms();

    conn.execute(
        "INSERT INTO sync_queue (action, method, params, payload, created_at) VALUES (?1, ?2, ?3, ?4, ?5)",
        params![action, method_value, params_json, payload_json, created_at],
    )
    .map_err(|e| e.to_string())?;
    Ok(conn.last_insert_rowid())
}

#[tauri::command]
fn queue_list(state: tauri::State<DbState>, limit: Option<u32>) -> Result<Vec<SyncJob>, String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let limit_value = limit.unwrap_or(200);
    let mut stmt = conn
        .prepare(
            "SELECT id, action, method, params, payload, created_at FROM sync_queue ORDER BY id ASC LIMIT ?1",
        )
        .map_err(|e| e.to_string())?;
    let rows = stmt
        .query_map(params![limit_value], |row| {
            let params_raw: String = row.get(3)?;
            let payload_raw: String = row.get(4)?;
            let params_value = serde_json::from_str(&params_raw).unwrap_or(serde_json::Value::Null);
            let payload_value =
                serde_json::from_str(&payload_raw).unwrap_or(serde_json::Value::Null);
            Ok(SyncJob {
                id: row.get(0)?,
                action: row.get(1)?,
                method: row.get(2)?,
                params: params_value,
                payload: payload_value,
                created_at: row.get(5)?,
            })
        })
        .map_err(|e| e.to_string())?;
    let mut jobs = Vec::new();
    for job in rows {
        jobs.push(job.map_err(|e| e.to_string())?);
    }
    Ok(jobs)
}

#[tauri::command]
fn queue_delete(state: tauri::State<DbState>, id: i64) -> Result<(), String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    conn.execute("DELETE FROM sync_queue WHERE id = ?1", params![id])
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn queue_clear(state: tauri::State<DbState>) -> Result<(), String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    conn.execute("DELETE FROM sync_queue", [])
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn queue_count(state: tauri::State<DbState>) -> Result<i64, String> {
    let conn = state
        .conn
        .lock()
        .map_err(|_| "Database lock poisoned".to_string())?;
    let count: i64 = conn
        .query_row("SELECT COUNT(*) FROM sync_queue", [], |row| row.get(0))
        .map_err(|e| e.to_string())?;
    Ok(count)
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
fn upload_tenant_document(
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
        "action": "uploadTenantDocument",
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

#[tauri::command]
fn download_file_to_path(
    app: tauri::AppHandle,
    url: String,
    file_path: String,
    download_id: String,
) -> Result<serde_json::Value, String> {
    if url.trim().is_empty() {
        return Err("Missing URL".to_string());
    }
    if file_path.trim().is_empty() {
        return Err("Missing file path".to_string());
    }

    let response = ureq::get(&url).call().map_err(|e| e.to_string())?;
    let total = response
        .header("Content-Length")
        .and_then(|val| val.parse::<u64>().ok())
        .unwrap_or(0);

    let mut reader = response.into_reader();
    let path = PathBuf::from(&file_path);
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    let mut file = std::fs::File::create(&file_path).map_err(|e| e.to_string())?;
    let mut buf = [0u8; 64 * 1024];
    let mut loaded = 0u64;
    let mut last_emit = 0u64;

    loop {
        let read = reader.read(&mut buf).map_err(|e| e.to_string())?;
        if read == 0 {
            let payload = DownloadProgress {
                download_id: download_id.clone(),
                loaded,
                total,
                done: true,
            };
            let _ = app.emit("download-progress", payload);
            break;
        }
        file.write_all(&buf[..read]).map_err(|e| e.to_string())?;
        loaded = loaded.saturating_add(read as u64);
        if loaded - last_emit >= 64 * 1024 || (total > 0 && loaded >= total) {
            let payload = DownloadProgress {
                download_id: download_id.clone(),
                loaded,
                total,
                done: loaded >= total && total > 0,
            };
            let _ = app.emit("download-progress", payload);
            last_emit = loaded;
        }
    }

    Ok(serde_json::json!({ "ok": true, "filePath": file_path }))
}

#[tauri::command]
fn copy_file_to_clipboard(file_path: String) -> Result<serde_json::Value, String> {
    if file_path.trim().is_empty() {
        return Err("Missing file path".to_string());
    }
    let path = PathBuf::from(&file_path);
    if !path.exists() {
        return Err("File not found".to_string());
    }

    #[cfg(windows)]
    {
        use std::ffi::OsStr;
        use std::os::windows::ffi::OsStrExt;
        use windows::Win32::Foundation::{GlobalFree, BOOL, HANDLE};
        use windows::Win32::System::DataExchange::{
            CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
        };
        use windows::Win32::System::Memory::{
            GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE, GMEM_ZEROINIT,
        };
        use windows::Win32::System::Ole::CF_HDROP;
        use windows::Win32::UI::Shell::DROPFILES;

        let mut wide: Vec<u16> = OsStr::new(&file_path).encode_wide().collect();
        wide.push(0);
        wide.push(0);
        let total_bytes = std::mem::size_of::<DROPFILES>() + wide.len() * 2;

        unsafe {
            let hmem = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, total_bytes)
                .map_err(|_| "Failed to allocate clipboard data".to_string())?;
            let locked = GlobalLock(hmem);
            if locked.is_null() {
                let _ = GlobalFree(hmem);
                return Err("Failed to lock clipboard data".to_string());
            }
            let dropfiles = locked as *mut DROPFILES;
            (*dropfiles).pFiles = std::mem::size_of::<DROPFILES>() as u32;
            (*dropfiles).fWide = BOOL(1);

            let list_ptr = (locked as *mut u8).add(std::mem::size_of::<DROPFILES>()) as *mut u16;
            std::ptr::copy_nonoverlapping(wide.as_ptr(), list_ptr, wide.len());

            let _ = GlobalUnlock(hmem);

            if OpenClipboard(None).is_err() {
                let _ = GlobalFree(hmem);
                return Err("Failed to open clipboard".to_string());
            }
            EmptyClipboard();
            if SetClipboardData(CF_HDROP.0 as u32, HANDLE(hmem.0 as isize)).is_err() {
                CloseClipboard();
                let _ = GlobalFree(hmem);
                return Err("Failed to set clipboard data".to_string());
            }
            CloseClipboard();
        }

        return Ok(serde_json::json!({ "ok": true }));
    }

    #[cfg(not(windows))]
    {
        let _ = path;
        Err("File copy is only supported on Windows.".to_string())
    }
}

#[tauri::command]
async fn open_whatsapp(app: tauri::AppHandle) -> Result<(), String> {
    use tauri::{WebviewUrl, WebviewWindowBuilder};

    let script = r##"
    (() => {
        "use strict";
        
        function getRawState() {
            // Strong indicators for Logged IN
            if (document.getElementById("pane-side")) return "IN";
            if (document.getElementById("main")) return "IN";
            if (document.querySelector("div[role='textbox']")) return "IN";
            
            // Strong indicators for Logged OUT (Login Screen)
            const bodyText = document.body.innerText || "";
            if (bodyText.includes("Use WhatsApp on your computer")) return "OUT";
            if (document.querySelector("canvas")) return "OUT"; // QR Code
            if (document.querySelector('[data-testid="qrcode"]')) return "OUT";
            if (document.querySelector('.landing-wrapper')) return "OUT";
            
            return "UNKNOWN";
        }

        function ensureOverlay() {
            let el = document.getElementById("wa-state-overlay");
            if (el) return el;
            el = document.createElement("div");
            el.id = "wa-state-overlay";
            el.style.cssText = "position:fixed;top:10px;right:10px;z-index:999999;background:rgba(0,0,0,0.8);color:white;padding:5px 10px;border-radius:4px;font-size:12px;pointer-events:none;font-weight:bold;";
            (document.body || document.documentElement).appendChild(el);
            return el;
        }

        setInterval(() => {
            const state = getRawState();
            
            // Use URL Hash for signaling (Title is prone to race conditions with WA)
            if (state !== "UNKNOWN") {
                const targetHash = "#wa_state=" + state;
                if (window.location.hash !== targetHash) {
                    window.location.hash = targetHash;
                }
            }

            // Visual Overlay
            const el = ensureOverlay();
            if (state === "IN") {
                el.innerText = "✅ Logged In";
                el.style.background = "rgba(16, 185, 129, 0.9)"; // Green
            } else if (state === "OUT") { 
                el.innerText = "🔒 Logged Out (QR)";
                el.style.background = "rgba(244, 63, 94, 0.9)"; // Red
            } else {
                el.innerText = "⏳ Detecting...";
                el.style.background = "rgba(0, 0, 0, 0.8)";
            }

        }, 200);
    })();
    "##;

    let win_builder = WebviewWindowBuilder::new(
        &app,
        "whatsapp_login",
        WebviewUrl::External("https://web.whatsapp.com".parse().unwrap()),
    )
    .title("WhatsApp Login")
    .inner_size(1000.0, 700.0)
    .initialization_script(script);

    let window = win_builder.build().map_err(|e| e.to_string())?;

    let win_clone = window.clone();
    let app_clone = app.clone();

    std::thread::spawn(move || {
        let mut last_processed_state = String::from("UNKNOWN");

        loop {
            std::thread::sleep(std::time::Duration::from_millis(500));

            // Use URL for reliability
            let url = match win_clone.url() {
                Ok(u) => u,
                Err(_) => break, // Window closed
            };
            let url_str = url.as_str();

            let current_state = if url_str.contains("#wa_state=IN") {
                "IN"
            } else if url_str.contains("#wa_state=OUT") {
                "OUT"
            } else {
                "UNKNOWN"
            };

            if current_state == "UNKNOWN" {
                continue;
            }

            if current_state != last_processed_state {
                if current_state == "IN" {
                    println!("Rust: Login Detected (Hash)");
                    let _ = app_clone.emit("whatsapp-login-success", ());

                    // Allow event propagation
                    std::thread::sleep(std::time::Duration::from_millis(1000));

                    // Close window instantly on login
                    let _ = win_clone.close();
                    break;
                } else if current_state == "OUT" {
                    println!("Rust: Logout Detected (Hash)");
                    let _ = app_clone.emit("whatsapp-logout", ());
                }
                last_processed_state = current_state.to_string();
            }
        }
    });

    Ok(())
}

#[tauri::command]
async fn send_whatsapp_message(
    app: tauri::AppHandle,
    phone: String,
    message: String,
    progress_label: String,
) -> Result<(), String> {
    use std::time::{Duration, Instant};
    use tauri::{WebviewUrl, WebviewWindowBuilder};

    // Script to Auto-Send and Show Progress
    // 1. Inject Overlay with Progress Label
    // 2. Wait for Send button
    // 3. Click it
    // 4. Mark title as WA_MSG_SENT_SUCCESS
    let script = format!(
        r##"
    (() => {{
        "use strict";
        let sent = false;
        let sendTriggered = false;
        let sendStartedAt = 0;
        let composerText = "";
        
        function ensureOverlay() {{
            let el = document.getElementById("wa-send-overlay");
            if (el) return el;
            
            // Add spinner keyframes
            if (!document.getElementById('wa-spinner-style')) {{
                const style = document.createElement('style');
                style.id = 'wa-spinner-style';
                style.textContent = '@keyframes wa-spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}';
                document.head.appendChild(style);
            }}
            
            el = document.createElement("div");
            el.id = "wa-send-overlay";
            el.style.cssText = "position:fixed;top:10px;right:10px;z-index:99999;background:linear-gradient(135deg, #0f2027, #2c5364);color:#fff;padding:8px 16px;border-radius:6px;font:bold 12px sans-serif;box-shadow:0 10px 30px rgba(0,0,0,0.5);border:1px solid rgba(255,255,255,0.2);display:flex;align-items:center;gap:8px;";
            el.innerHTML = '<span class="wa-spinner" style="width:14px;height:14px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:wa-spin 0.8s linear infinite;"></span><span class="wa-text">{}</span>'; 
            document.body.appendChild(el);
            return el;
        }}

        function findSendButton() {{
             return document.querySelector('span[data-icon="send"]') || 
                    document.querySelector('button[aria-label="Send"]');
        }}

        function getComposer() {{
            return (
                document.querySelector('[data-testid="conversation-compose-box-input"]') ||
                document.querySelector('div[contenteditable="true"][data-tab="10"]') ||
                document.querySelector('div[contenteditable="true"][data-tab]')
            );
        }}

        function getComposerText() {{
            const el = getComposer();
            if (!el) return "";
            return (el.innerText || el.textContent || "").trim();
        }}

        function getRawState() {{
            if (document.getElementById("pane-side") || document.getElementById("main")) return "IN";
            const body = document.body.innerText || "";
            if (body.includes("Use WhatsApp on your computer") || document.querySelector("canvas")) return "OUT";
            return "UNKNOWN";
        }}

        const interval = setInterval(() => {{
            const el = ensureOverlay();
            
            if (sent) {{
                // Close instantly without notification
                clearInterval(interval);
                return;
            }}

            const state = getRawState();
            
            // Communicate State via Hash
            if (state === "OUT") {{
                 const textSpan = el.querySelector('.wa-text');
                 if (textSpan) textSpan.innerText = "Logged Out! Please login on this screen.";
                 el.style.background = "linear-gradient(to right, #cb2d3e, #ef473a)";
                 if (!window.location.hash.includes("wa_state=OUT")) {{
                     window.location.hash = "wa_state=OUT";
                 }}
            }} else if (state === "IN") {{
                 if (!window.location.hash.includes("wa_state=IN") && !window.location.hash.includes("wa_msg_sent")) {{
                     window.location.hash = "wa_state=IN";
                 }}
                 const textSpan = el.querySelector('.wa-text');
                 if (textSpan) textSpan.innerText = "{}"; // Restore original progress label
                 el.style.background = "linear-gradient(135deg, #0f2027, #2c5364)";

                 // Attempt Send
                 if (!sendTriggered) {{
                     const btn = findSendButton();
                     if (btn) {{
                        composerText = getComposerText();
                        sendTriggered = true;
                        sendStartedAt = Date.now();
                        btn.click();
                     }}
                 }} else if (!sent) {{
                     const currentText = getComposerText();
                     const elapsed = Date.now() - sendStartedAt;
                     const cleared = composerText
                        ? currentText === ""
                        : elapsed > 3000 && currentText === "";
                     if (cleared || elapsed > 20000) {{
                        sent = true;
                        setTimeout(() => {{
                            document.title = "WA_MSG_SENT_SUCCESS";
                            window.location.hash = "wa_msg_sent_success";
                        }}, 1500);
                     }}
                 }}
            }}
        }}, 1000);

        setTimeout(() => {{
            if (!sent) {{
                document.title = "WA_MSG_SENT_TIMEOUT";
            }}
        }}, 600000); // 10 minute timeout to allow for login
    }})();
    "##,
        progress_label, progress_label
    );

    let url_str = format!(
        "https://web.whatsapp.com/send?phone={}&text={}&app_absent=0",
        phone,
        urlencoding::encode(&message)
    );

    let win_builder = WebviewWindowBuilder::new(
        &app,
        format!(
            "whatsapp_send_{}",
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .as_millis()
        ),
        WebviewUrl::External(url_str.parse().unwrap()),
    )
    .title(format!("Sending: {}", progress_label))
    .inner_size(800.0, 600.0)
    .initialization_script(&script);

    let window = win_builder.build().map_err(|e| e.to_string())?;

    // Block and poll for success directly
    let start = Instant::now();
    let mut success = false;
    let mut last_send_state = String::from("UNKNOWN");
    let app_clone = app.clone();

    loop {
        tokio::time::sleep(Duration::from_millis(500)).await;

        if start.elapsed().as_secs() > 600 {
            // 10 minutes matches JS
            let _ = window.close();
            break;
        }

        if let Ok(url) = window.url() {
            let url_str = url.as_str();

            // Check Success
            if url_str.contains("wa_msg_sent_success") {
                success = true;
            }

            // Check Login State Changes
            let current_send_state = if url_str.contains("wa_state=OUT") {
                "OUT"
            } else if url_str.contains("wa_state=IN") {
                "IN"
            } else {
                "UNKNOWN"
            };

            if current_send_state != "UNKNOWN" && current_send_state != last_send_state {
                if current_send_state == "OUT" {
                    println!("Rust (Send): Logout Detected");
                    let _ = app_clone.emit("whatsapp-logout", ());
                } else if current_send_state == "IN" {
                    println!("Rust (Send): Login Detected");
                    let _ = app_clone.emit("whatsapp-login-success", ());
                }
                last_send_state = current_send_state.to_string();
            }
        } else {
            // Window closed
            break;
        }

        // Double check title
        if !success {
            if let Ok(title) = window.title() {
                if title.contains("WA_MSG_SENT_SUCCESS") {
                    success = true;
                }
            }
        }

        if success {
            // Close instantly without delay
            let _ = window.close();
            return Ok(());
        }
    }

    if !success {
        return Err("Timeout or Window Closed".to_string());
    }

    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .manage(UploadState::default())
        .setup(|app| {
            let conn =
                setup_db(&app.handle()).map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
            app.manage(DbState {
                conn: Mutex::new(conn),
            });
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            cache_get,
            cache_set,
            cache_set_many,
            cache_delete,
            cache_delete_prefix,
            queue_add,
            queue_list,
            queue_delete,
            queue_clear,
            queue_count,
            upload_payment_attachment,
            upload_tenant_document,
            cancel_upload,
            download_file_to_path,
            copy_file_to_clipboard,
            ensure_temp_dir,
            open_whatsapp,
            send_whatsapp_message
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
