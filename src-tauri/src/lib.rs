use rusqlite::{params, Connection};
use serde::Serialize;
use std::collections::HashMap;
use std::io::{self, Read};
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
fn cancel_upload(state: tauri::State<UploadState>, upload_id: String) -> bool {
    state.cancel(&upload_id)
}

#[tauri::command]
async fn open_whatsapp(app: tauri::AppHandle) -> Result<(), String> {
    use tauri::{WebviewUrl, WebviewWindowBuilder};

    let script = r##"
    (() => {
        "use strict";
        const STATE = { IN: "IN", OUT: "OUT", UNKNOWN: "UNKNOWN" };
        let isSuccessHandled = false;

        function detect() {
            const paneSide = document.getElementById("pane-side");
            const main = document.getElementById("main");
            if (paneSide || main) return STATE.IN;
            const app = document.getElementById("app") || document.body;
            const qrCanvas = app ? app.querySelector("div[data-ref] canvas, canvas") : null;
            if (qrCanvas) return STATE.OUT;
            return STATE.UNKNOWN;
        }

        function ensureOverlay() {
            let el = document.getElementById("wa-login-overlay");
            if (el) return el;
            el = document.createElement("div");
            el.id = "wa-login-overlay";
            // Default small overlay style
            el.style.cssText = "position:fixed;top:12px;right:12px;z-index:2147483647;background:#111;color:#fff;padding:10px 12px;border-radius:12px;font:13px/1.35 system-ui,sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.35);border:1px solid rgba(255,255,255,.18);opacity:.95;pointer-events:none;transition:all 0.5s ease;";
            el.innerHTML = '<div style="font-weight:700;margin-bottom:4px;">WhatsApp Web</div><div id="wa-login-status" style="font-weight:600;">Starting…</div>';
            (document.body || document.documentElement).appendChild(el);
            return el;
        }

        function setStatus(state) {
            if (isSuccessHandled) return;

            ensureOverlay();
            const s = document.getElementById("wa-login-status");
            if (!s) return;

            if (state === STATE.IN) {
                isSuccessHandled = true;
                s.textContent = "✅ Logged in";
                
                // Signal backend to close immediately
                document.title = "WA_LOGGED_IN_SUCCESS";
                if (window.location.hash !== "#wa_login_success") {
                    window.location.hash = "wa_login_success";
                }

            } else if (state === STATE.OUT) {
                s.textContent = "🔒 Logged out (QR screen)";
            } else {
                s.textContent = "⏳ Loading / Unknown";
            }
        }

        function tick() {
            const state = detect();
            setStatus(state);
        }
        setInterval(tick, 1000);
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

    // Spawn a polling task to check for the title change
    let win_clone = window.clone();
    let app_clone = app.clone();

    std::thread::spawn(move || {
        loop {
            std::thread::sleep(std::time::Duration::from_millis(500));
            let mut detected = false;

            // Check Title
            if let Ok(title) = win_clone.title() {
                if title.contains("WA_LOGGED_IN_SUCCESS") {
                    println!("WA Login Detected via Title");
                    detected = true;
                }
            } else {
                break; // Window closed
            }

            // Check URL Hash (fallback)
            if !detected {
                if let Ok(url) = win_clone.url() {
                    if url.as_str().contains("wa_login_success") {
                        println!("WA Login Detected via URL Hash");
                        detected = true;
                    }
                }
            }

            if detected {
                let _ = app_clone.emit("whatsapp-login-success", ());
                let _ = win_clone.close();
                break;
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
        
        function ensureOverlay() {{
            let el = document.getElementById("wa-send-overlay");
            if (el) return el;
            el = document.createElement("div");
            el.id = "wa-send-overlay";
            el.style.cssText = "position:fixed;top:10px;right:10px;z-index:99999;background:linear-gradient(135deg, #0f2027, #203a43, #2c5364);color:#fff;padding:8px 16px;border-radius:20px;font:bold 12px sans-serif;box-shadow:0 4px 15px rgba(0,0,0,0.5);border:1px solid rgba(255,255,255,0.2);";
            el.innerText = "{}"; 
            document.body.appendChild(el);
            return el;
        }}

        // Helper to find button by aria-label or data-icon
        function findSendButton() {{
            return document.querySelector('span[data-icon="send"]') || 
                   document.querySelector('button[aria-label="Send"]');
        }}

        const interval = setInterval(() => {{
            ensureOverlay();
            
            if (sent) {{
                clearInterval(interval);
                return;
            }}

            const btn = findSendButton();
            if (btn) {{
                 // Simulate click
                const event = new MouseEvent('click', {{
                    view: window,
                    bubbles: true,
                    cancelable: true
                }});
                btn.dispatchEvent(event);
                btn.click();
                
                sent = true;
                
                // Allow some time for network request
                setTimeout(() => {{
                    document.title = "WA_MSG_SENT_SUCCESS";
                    window.location.hash = "wa_msg_sent_success";
                }}, 3000); // 3s wait after click
            }}
        }}, 1000);

        // Safety timeout 45s
        setTimeout(() => {{
            if (!sent) {{
                document.title = "WA_MSG_SENT_TIMEOUT";
            }}
        }}, 45000);
    }})();
    "##,
        progress_label
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

    // Block and poll for success directly in this async function behavior
    // We use tokio::time::sleep to yield but keep this task alive
    let start = Instant::now();
    let mut success = false;

    loop {
        // Yield to event loop
        tokio::time::sleep(Duration::from_millis(500)).await;

        if start.elapsed().as_secs() > 60 {
            let _ = window.close();
            break;
        }

        // We must run window interactions on valid thread/context or just check normally.
        // tauri::Window is thread-safe cloneable.

        if let Ok(title) = window.title() {
            if title.contains("WA_MSG_SENT_SUCCESS") {
                success = true;
            }
        } else {
            // Window closed manually by user or error
            break;
        }

        if !success {
            if let Ok(url) = window.url() {
                if url.as_str().contains("wa_msg_sent_success") {
                    success = true;
                }
            }
        }

        if success {
            // Wait small buffer
            tokio::time::sleep(Duration::from_millis(1000)).await;
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
            cache_delete,
            cache_delete_prefix,
            queue_add,
            queue_list,
            queue_delete,
            queue_clear,
            queue_count,
            upload_payment_attachment,
            cancel_upload,
            open_whatsapp,
            send_whatsapp_message
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
