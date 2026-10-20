//! Распознавание формата по первым байтам и содержимому архива.
//!
//! Файлы из мессенджеров часто приходят без расширения или с неверным
//! (например, `.doc`, который на самом деле `.docx`), поэтому смотрим внутрь.

use std::io::{Cursor, Read};

/// Что лежит в файле. Названия совпадают с теми, что понимает страница просмотра.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Pdf,
    /// Word 2007+ (.docx)
    Docx,
    /// Word 97–2003 (.doc)
    Doc,
    /// Любая таблица: .xlsx, .xls, .ods
    Sheet,
    /// PowerPoint 2007+ (.pptx)
    Pptx,
    /// PowerPoint 97–2003 (.ppt), показать не можем
    PptOld,
    /// Текстовый документ OpenDocument (.odt), показать не можем
    Odt,
    /// Офисный файл, зашифрованный паролем
    Encrypted,
    Rtf,
    Image,
    Video,
    Audio,
    Text,
    Unknown,
}

impl Kind {
    pub fn as_str(self) -> &'static str {
        match self {
            Kind::Pdf => "pdf",
            Kind::Docx => "docx",
            Kind::Doc => "doc",
            Kind::Sheet => "sheet",
            Kind::Pptx => "pptx",
            Kind::PptOld => "ppt-old",
            Kind::Odt => "odt",
            Kind::Encrypted => "encrypted",
            Kind::Rtf => "rtf",
            Kind::Image => "image",
            Kind::Video => "video",
            Kind::Audio => "audio",
            Kind::Text => "text",
            Kind::Unknown => "unknown",
        }
    }
}

pub fn detect(data: &[u8]) -> Kind {
    let at = |offset: usize, sig: &[u8]| data.get(offset..offset + sig.len()) == Some(sig);

    if at(0, b"%PDF") {
        return Kind::Pdf;
    }
    if at(0, b"PK\x03\x04") {
        return detect_zip(data);
    }
    if at(0, &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) {
        return detect_ole(data);
    }
    if at(0, b"{\\rtf") {
        return Kind::Rtf;
    }
    if at(0, b"\x89PNG") || at(0, &[0xFF, 0xD8, 0xFF]) || at(0, b"GIF8") || at(0, b"BM") && data.len() > 26 {
        return Kind::Image;
    }
    if at(0, b"RIFF") {
        if at(8, b"WEBP") {
            return Kind::Image;
        }
        if at(8, b"WAVE") {
            return Kind::Audio;
        }
        if at(8, b"AVI ") {
            return Kind::Video;
        }
    }
    if at(4, b"ftyp") {
        // ISO-контейнер: и видео, и HEIC-фото, и m4a
        let brand = data.get(8..12).unwrap_or_default();
        return match brand {
            b"heic" | b"heix" | b"mif1" | b"msf1" | b"avif" => Kind::Image,
            b"M4A " | b"M4B " => Kind::Audio,
            _ => Kind::Video,
        };
    }
    if at(0, &[0x1A, 0x45, 0xDF, 0xA3]) {
        return Kind::Video; // Matroska / WebM
    }
    if at(0, b"ID3") || at(0, b"OggS") || at(0, b"fLaC") || (data.len() > 2 && data[0] == 0xFF && data[1] & 0xE0 == 0xE0) {
        return Kind::Audio;
    }
    if looks_like_text(data) {
        return Kind::Text;
    }
    Kind::Unknown
}

fn detect_zip(data: &[u8]) -> Kind {
    let Ok(mut zip) = zip::ZipArchive::new(Cursor::new(data)) else {
        return Kind::Unknown;
    };
    let has = |zip: &zip::ZipArchive<_>, name: &str| zip.index_for_name(name).is_some();
    if has(&zip, "word/document.xml") {
        return Kind::Docx;
    }
    if has(&zip, "xl/workbook.xml") {
        return Kind::Sheet;
    }
    if has(&zip, "ppt/presentation.xml") {
        return Kind::Pptx;
    }
    // OpenDocument пишет свой тип в файл `mimetype`
    let mut mime = String::new();
    if let Ok(mut f) = zip.by_name("mimetype") {
        let _ = f.by_ref().take(100).read_to_string(&mut mime);
    }
    match mime.trim() {
        "application/vnd.oasis.opendocument.spreadsheet" => Kind::Sheet,
        "application/vnd.oasis.opendocument.text" => Kind::Odt,
        _ => Kind::Unknown,
    }
}

fn detect_ole(data: &[u8]) -> Kind {
    let Ok(file) = cfb::CompoundFile::open(Cursor::new(data)) else {
        return Kind::Unknown;
    };
    if file.is_stream("/EncryptedPackage") {
        return Kind::Encrypted;
    }
    if file.is_stream("/WordDocument") {
        return Kind::Doc;
    }
    if file.is_stream("/Workbook") || file.is_stream("/Book") {
        return Kind::Sheet;
    }
    if file.is_stream("/PowerPoint Document") {
        return Kind::PptOld;
    }
    Kind::Unknown
}

/// Текст, если начало файла — корректный UTF-8 (или UTF-16 с меткой) без управляющих символов.
fn looks_like_text(data: &[u8]) -> bool {
    if data.is_empty() {
        return false;
    }
    if data.starts_with(&[0xFF, 0xFE]) || data.starts_with(&[0xFE, 0xFF]) {
        return true;
    }
    let head = &data[..data.len().min(8192)];
    let text = match std::str::from_utf8(head) {
        Ok(t) => t,
        // обрезали посреди многобайтового символа
        Err(e) if e.error_len().is_none() => std::str::from_utf8(&head[..e.valid_up_to()]).unwrap_or_default(),
        Err(_) => return looks_like_single_byte_text(head),
    };
    let bad = text.chars().filter(|c| c.is_control() && !matches!(c, '\n' | '\r' | '\t' | '\x0C')).count();
    bad * 100 <= text.chars().count()
}

/// Однобайтовые кодировки вроде Windows-1251: почти нет байтов меньше 0x20.
fn looks_like_single_byte_text(head: &[u8]) -> bool {
    let bad = head.iter().filter(|&&b| b < 0x20 && !matches!(b, b'\n' | b'\r' | b'\t')).count();
    bad == 0
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn signatures() {
        assert_eq!(detect(b"%PDF-1.7\n..."), Kind::Pdf);
        assert_eq!(detect(b"\x89PNG\r\n\x1a\n...."), Kind::Image);
        assert_eq!(detect(b"{\\rtf1\\ansi ..."), Kind::Rtf);
        assert_eq!(detect(b"\0\0\0\x18ftypheic\0\0\0\0"), Kind::Image);
        assert_eq!(detect(b"\0\0\0\x18ftypisom\0\0\0\0"), Kind::Video);
        assert_eq!(detect(b"ID3\x04\0\0\0\0\0\0"), Kind::Audio);
    }

    #[test]
    fn text() {
        assert_eq!(detect("Привет, это обычный текст\nвторая строка".as_bytes()), Kind::Text);
        // Windows-1251: «Привет»
        assert_eq!(detect(&[0xCF, 0xF0, 0xE8, 0xE2, 0xE5, 0xF2, b'\n']), Kind::Text);
        assert_eq!(detect(&[0, 1, 2, 3, 4, 5, 6, 7, 0xFE, 0x99]), Kind::Unknown);
        assert_eq!(detect(b""), Kind::Unknown);
    }
}
