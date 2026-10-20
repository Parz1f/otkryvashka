//! Логика Открывашки, которой не нужен ни Android, ни браузер:
//! распознавание формата файла по содержимому и чтение текста из `.doc`.

mod detect;
mod doc;

pub use detect::{detect, Kind};
pub use doc::{doc_text, DocError};
