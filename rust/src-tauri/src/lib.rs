//! Оболочка Открывашки на Tauri. Документы показывает та же страница, что и в
//! обычной версии, а Rust берёт на себя работу с файлами:
//! - распознаёт формат по содержимому (`detect_format`);
//! - читает текст старых документов Word (`doc_text`);
//! - отдаёт странице файл, открытый из другого приложения (`incoming_info`, `take_incoming`);
//! - закрывает приложение по кнопке «Назад» на главном экране (`quit`).

use std::borrow::Cow;
use std::fs;
use std::path::PathBuf;

use serde::{Deserialize, Serialize};
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{AppHandle, Manager};

/// Файл из «Открыть с помощью». Его кладёт в кэш приложения MainActivity.kt:
/// сначала байты в `incoming/file`, потом описание в `incoming/meta.json`.
#[derive(Serialize, Deserialize)]
struct Incoming {
    name: String,
    mime: String,
    size: u64,
    /// Формат, распознанный по содержимому; заполняет `incoming_info`
    #[serde(default)]
    kind: String,
}

fn incoming_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let cache = app.path().app_cache_dir().map_err(|e| e.to_string())?;
    Ok(cache.join("incoming"))
}

/// Байты, которые страница передала в команду как ArrayBuffer.
/// На компьютере они приходят как есть, а на Android — JSON-массивом чисел.
fn raw_body<'a>(request: &'a Request<'_>) -> Result<Cow<'a, [u8]>, String> {
    match request.body() {
        InvokeBody::Raw(bytes) => Ok(Cow::Borrowed(bytes)),
        InvokeBody::Json(value) => Vec::<u8>::deserialize(value)
            .map(Cow::Owned)
            .map_err(|_| "Команде нужно передать файл как ArrayBuffer.".to_string()),
    }
}

#[tauri::command]
fn detect_format(request: Request<'_>) -> Result<&'static str, String> {
    Ok(otkryvashka_core::detect(&raw_body(&request)?).as_str())
}

#[tauri::command]
fn doc_text(request: Request<'_>) -> Result<String, String> {
    otkryvashka_core::doc_text(&raw_body(&request)?).map_err(|e| e.to_string())
}

#[tauri::command]
fn incoming_info(app: AppHandle) -> Result<Option<Incoming>, String> {
    let dir = incoming_dir(&app)?;
    let Ok(json) = fs::read_to_string(dir.join("meta.json")) else {
        return Ok(None);
    };
    let mut info: Incoming = serde_json::from_str(&json).map_err(|e| e.to_string())?;
    if let Ok(bytes) = fs::read(dir.join("file")) {
        info.kind = otkryvashka_core::detect(&bytes).as_str().to_string();
    }
    Ok(Some(info))
}

/// Отдаёт файл и убирает его из кэша, чтобы он не открылся второй раз.
#[tauri::command]
fn take_incoming(app: AppHandle) -> Result<Response, String> {
    let dir = incoming_dir(&app)?;
    let bytes = fs::read(dir.join("file")).map_err(|_| "Файл уже открыт или недоступен.".to_string())?;
    let _ = fs::remove_file(dir.join("meta.json"));
    let _ = fs::remove_file(dir.join("file"));
    Ok(Response::new(bytes))
}

#[tauri::command]
fn quit(app: AppHandle) {
    app.exit(0);
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![detect_format, doc_text, incoming_info, take_incoming, quit])
        .run(tauri::generate_context!())
        .expect("не удалось запустить Открывашку");
}
