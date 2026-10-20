//! Текст из документов Word 97–2003 (.doc).
//!
//! Такой файл — контейнер OLE (CFB) с потоками `WordDocument` и `0Table`/`1Table`.
//! Текст хранится кусками («pieces»); где лежит каждый кусок и в какой он кодировке,
//! описывает таблица кусков (Clx/PlcPcd) в табличном потоке. Формат описан
//! в спецификации Microsoft [MS-DOC].

use std::fmt;
use std::io::{Cursor, Read};

#[derive(Debug, PartialEq, Eq)]
pub enum DocError {
    /// Это не документ Word 97–2003
    NotWord,
    /// Документ защищён паролем
    Encrypted,
    /// Внутренние таблицы документа повреждены
    Corrupt(&'static str),
}

impl fmt::Display for DocError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            DocError::NotWord => write!(f, "Это не документ Word 97–2003."),
            DocError::Encrypted => write!(f, "Документ защищён паролем."),
            DocError::Corrupt(what) => write!(f, "Документ повреждён ({what})."),
        }
    }
}

impl std::error::Error for DocError {}

// Смещения полей в FIB (заголовке потока WordDocument), [MS-DOC] 2.5.1–2.5.6
const FIB_IDENT: usize = 0x0000;
const FIB_FLAGS: usize = 0x000A;
const FIB_CCP_TEXT: usize = 0x004C;
const FIB_FC_CLX: usize = 0x01A2;
const FIB_LCB_CLX: usize = 0x01A6;
const WORD_IDENT: u16 = 0xA5EC;
const FLAG_ENCRYPTED: u16 = 0x0100;
const FLAG_TABLE1: u16 = 0x0200;

/// Текст основной части документа, без колонтитулов и сносок.
pub fn doc_text(data: &[u8]) -> Result<String, DocError> {
    let mut file = cfb::CompoundFile::open(Cursor::new(data)).map_err(|_| DocError::NotWord)?;
    let word = read_stream(&mut file, "/WordDocument").ok_or(DocError::NotWord)?;

    if u16_at(&word, FIB_IDENT) != Some(WORD_IDENT) {
        return Err(DocError::NotWord);
    }
    let flags = u16_at(&word, FIB_FLAGS).ok_or(DocError::Corrupt("заголовок"))?;
    if flags & FLAG_ENCRYPTED != 0 {
        return Err(DocError::Encrypted);
    }
    let table_name = if flags & FLAG_TABLE1 != 0 { "/1Table" } else { "/0Table" };
    let table = read_stream(&mut file, table_name).ok_or(DocError::Corrupt("нет таблицы"))?;

    let ccp_text = u32_at(&word, FIB_CCP_TEXT).ok_or(DocError::Corrupt("заголовок"))? as usize;
    let fc_clx = u32_at(&word, FIB_FC_CLX).ok_or(DocError::Corrupt("заголовок"))? as usize;
    let lcb_clx = u32_at(&word, FIB_LCB_CLX).ok_or(DocError::Corrupt("заголовок"))? as usize;
    let clx = table.get(fc_clx..fc_clx + lcb_clx).ok_or(DocError::Corrupt("таблица кусков"))?;

    let pieces = parse_clx(clx)?;
    let mut raw = String::new();
    for p in pieces {
        if p.cp_start >= ccp_text {
            break;
        }
        let len = p.cp_end.min(ccp_text) - p.cp_start;
        if p.compressed {
            let bytes = word.get(p.fc..p.fc + len).ok_or(DocError::Corrupt("кусок текста"))?;
            raw.extend(bytes.iter().map(|&b| cp1252(b)));
        } else {
            let bytes = word.get(p.fc..p.fc + len * 2).ok_or(DocError::Corrupt("кусок текста"))?;
            let units: Vec<u16> = bytes.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
            raw.extend(char::decode_utf16(units).map(|r| r.unwrap_or('\u{FFFD}')));
        }
    }
    Ok(clean(&raw))
}

struct Piece {
    cp_start: usize,
    cp_end: usize,
    /// Смещение в потоке WordDocument, в байтах
    fc: usize,
    /// true — по байту на символ (Windows-1252), false — UTF-16LE
    compressed: bool,
}

/// Clx: сначала любое число блоков Prc (0x01), затем один Pcdt (0x02) с таблицей кусков.
fn parse_clx(clx: &[u8]) -> Result<Vec<Piece>, DocError> {
    let mut pos = 0;
    while clx.get(pos) == Some(&0x01) {
        let size = u16_at(clx, pos + 1).ok_or(DocError::Corrupt("таблица кусков"))? as usize;
        pos += 3 + size;
    }
    if clx.get(pos) != Some(&0x02) {
        return Err(DocError::Corrupt("таблица кусков"));
    }
    let lcb = u32_at(clx, pos + 1).ok_or(DocError::Corrupt("таблица кусков"))? as usize;
    let plc = clx.get(pos + 5..pos + 5 + lcb).ok_or(DocError::Corrupt("таблица кусков"))?;

    // PlcPcd: n+1 позиций символов (по 4 байта), затем n описаний кусков (по 8 байт)
    if lcb < 4 || (lcb - 4) % 12 != 0 {
        return Err(DocError::Corrupt("таблица кусков"));
    }
    let n = (lcb - 4) / 12;
    let cp = |i: usize| u32_at(plc, i * 4).unwrap_or(0) as usize;
    let mut pieces = Vec::with_capacity(n);
    for i in 0..n {
        let fc_raw = u32_at(plc, (n + 1) * 4 + i * 8 + 2).unwrap_or(0);
        let compressed = fc_raw & 0x4000_0000 != 0;
        let fc = (fc_raw & 0x3FFF_FFFF) as usize;
        let (cp_start, cp_end) = (cp(i), cp(i + 1));
        if cp_end < cp_start {
            return Err(DocError::Corrupt("таблица кусков"));
        }
        pieces.push(Piece { cp_start, cp_end, fc: if compressed { fc / 2 } else { fc }, compressed });
    }
    Ok(pieces)
}

/// Убирает служебные символы Word: коды полей, метки ячеек, объекты.
fn clean(raw: &str) -> String {
    let mut out = String::with_capacity(raw.len());
    // Для каждого открытого поля: дошли ли уже до его результата (после 0x14)
    let mut fields: Vec<bool> = Vec::new();
    let mut prev_cell_mark = false;
    for c in raw.chars() {
        let was_cell_mark = std::mem::replace(&mut prev_cell_mark, c == '\u{07}');
        match c {
            '\u{13}' => {
                fields.push(false);
                continue;
            }
            '\u{14}' => {
                if let Some(last) = fields.last_mut() {
                    *last = true;
                }
                continue;
            }
            '\u{15}' => {
                fields.pop();
                continue;
            }
            _ => {}
        }
        // Код поля (например, «PAGE» или «HYPERLINK "…"») не показываем, только его результат
        if fields.iter().any(|shown| !shown) {
            continue;
        }
        match c {
            '\r' | '\u{0B}' | '\u{0C}' => out.push('\n'),
            // Конец ячейки таблицы; две метки подряд — конец строки таблицы
            '\u{07}' if was_cell_mark => {
                out.pop();
                out.push('\n');
            }
            '\u{07}' => out.push('\t'),
            '\u{1E}' => out.push('-'),  // неразрывный дефис
            '\u{1F}' | '\u{01}' | '\u{08}' | '\u{02}' | '\u{05}' => {}
            c if c.is_control() && c != '\t' && c != '\n' => {}
            c => out.push(c),
        }
    }
    let lines: Vec<&str> = out.lines().map(|l| l.trim_end_matches('\t')).collect();
    let mut text = lines.join("\n");
    while text.contains("\n\n\n") {
        text = text.replace("\n\n\n", "\n\n");
    }
    text.trim().to_string()
}

fn read_stream(file: &mut cfb::CompoundFile<Cursor<&[u8]>>, name: &str) -> Option<Vec<u8>> {
    let mut stream = file.open_stream(name).ok()?;
    let mut buf = Vec::new();
    stream.read_to_end(&mut buf).ok()?;
    Some(buf)
}

fn u16_at(b: &[u8], i: usize) -> Option<u16> {
    Some(u16::from_le_bytes(b.get(i..i + 2)?.try_into().ok()?))
}

fn u32_at(b: &[u8], i: usize) -> Option<u32> {
    Some(u32::from_le_bytes(b.get(i..i + 4)?.try_into().ok()?))
}

/// Windows-1252: совпадает с Latin-1, кроме диапазона 0x80–0x9F.
fn cp1252(b: u8) -> char {
    const HIGH: [char; 32] = [
        '€', '\u{81}', '‚', 'ƒ', '„', '…', '†', '‡', 'ˆ', '‰', 'Š', '‹', 'Œ', '\u{8D}', 'Ž', '\u{8F}',
        '\u{90}', '‘', '’', '“', '”', '•', '–', '—', '˜', '™', 'š', '›', 'œ', '\u{9D}', 'ž', 'Ÿ',
    ];
    match b {
        0x80..=0x9F => HIGH[(b - 0x80) as usize],
        _ => b as char,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fields_show_only_their_result() {
        let raw = "Стр. \u{13} PAGE \u{14}3\u{15} из 5\r";
        assert_eq!(clean(raw), "Стр. 3 из 5");
    }

    #[test]
    fn nested_fields() {
        let raw = "А\u{13}IF \u{13}PAGE\u{14}1\u{15} = 1 \u{14}да\u{15}Б";
        assert_eq!(clean(raw), "АдаБ");
    }

    #[test]
    fn table_cells_become_tabs() {
        let raw = "Имя\u{07}Сумма\u{07}\u{07}Иван\u{07}100\u{07}\u{07}";
        assert_eq!(clean(raw), "Имя\tСумма\nИван\t100");
    }

    #[test]
    fn not_a_doc() {
        assert_eq!(doc_text(b"hello"), Err(DocError::NotWord));
    }
}
