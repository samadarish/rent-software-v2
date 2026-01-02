use serde::Serialize;
use tauri::{Emitter, Manager};
use rusqlite::{Connection, params};
use std::collections::HashMap;
use std::io::{self, Read};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};

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
    let dir = app
        .path()
        .app_data_dir()
        .map_err(|e| e.to_string())?;
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
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
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
fn cache_set(state: tauri::State<DbState>, key: String, value: serde_json::Value) -> Result<(), String> {
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
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
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
    conn.execute("DELETE FROM local_cache WHERE key = ?1", params![key])
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn cache_delete_prefix(state: tauri::State<DbState>, prefix: String) -> Result<(), String> {
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
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
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
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
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
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
            let payload_value = serde_json::from_str(&payload_raw).unwrap_or(serde_json::Value::Null);
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
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
    conn.execute("DELETE FROM sync_queue WHERE id = ?1", params![id])
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn queue_clear(state: tauri::State<DbState>) -> Result<(), String> {
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
    conn.execute("DELETE FROM sync_queue", [])
        .map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn queue_count(state: tauri::State<DbState>) -> Result<i64, String> {
    let conn = state.conn.lock().map_err(|_| "Database lock poisoned".to_string())?;
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

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_fs::init())
        .manage(UploadState::default())
        .setup(|app| {
            let conn = setup_db(&app.handle())
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
            app.manage(DbState { conn: Mutex::new(conn) });
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
            cancel_upload
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
