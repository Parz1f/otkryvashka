//! Проверки на настоящих файлах, сделанных LibreOffice, python-docx и python-pptx.

use otkryvashka_core::{detect, doc_text, Kind};

fn file(name: &str) -> Vec<u8> {
    std::fs::read(format!("{}/tests/files/{name}", env!("CARGO_MANIFEST_DIR"))).unwrap()
}

#[test]
fn detects_office_files_by_content() {
    assert_eq!(detect(&file("ru.doc")), Kind::Doc);
    assert_eq!(detect(&file("table.xlsx")), Kind::Sheet);
    assert_eq!(detect(&file("table.xls")), Kind::Sheet);
    assert_eq!(detect(&file("slides.pptx")), Kind::Pptx);
    assert_eq!(detect(&file("ru.odt")), Kind::Odt);
}

#[test]
fn reads_russian_doc() {
    let text = doc_text(&file("ru.doc")).unwrap();
    println!("{text}");
    assert!(text.starts_with("Ирвинг Фишер\n"));
    assert!(text.contains("Уравнение обмена: MV = PT. Цена — 1 500 ₽, «кавычки»."));
    assert!(text.contains("Имя\tСумма\nИван\t100"));
    assert!(text.ends_with("Последний абзац."));
}

#[test]
fn reads_latin_doc() {
    let text = doc_text(&file("latin.doc")).unwrap();
    assert_eq!(text, "Plain English only, no Cyrillic at all.\nSecond paragraph.");
}
