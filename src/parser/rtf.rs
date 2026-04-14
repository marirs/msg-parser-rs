//! RTF decompression per MS-OXRTFCP.
//!
//! The compressed format uses a simple LZ77 variant with a pre-filled
//! dictionary. See: <https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp>

/// Pre-filled dictionary used by the LZ algorithm.
const INIT_DICT: &[u8] = b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}\
{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArial\
Times New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \
\\pard\\plain\\f0\\fs20\\b\\i\\ul\\ob\\strike\\scaps\\shad\\outl\\pn\\v\\super\\sub\\nosupersub\
{\\*\\teletypertab{\\stylesheet{\\Normal;}}";

const INIT_DICT_LEN: usize = 207;
const COMP_MAGIC: u32 = 0x75465A4C; // "LZFu"
const UNCOMP_MAGIC: u32 = 0x414C454D; // "MELA"

/// Decompress RTF from raw bytes (the binary content of PidTagRtfCompressed).
///
/// Returns `None` if the input is too short or has an invalid header.
/// Returns the decompressed RTF as a byte vector.
pub(crate) fn decompress_rtf(data: &[u8]) -> Option<Vec<u8>> {
    if data.len() < 16 {
        return None;
    }

    let comp_size = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let raw_size = u32::from_le_bytes([data[4], data[5], data[6], data[7]]) as usize;
    let magic = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
    let _crc = u32::from_le_bytes([data[12], data[13], data[14], data[15]]);

    // Validate that we have enough data
    if comp_size + 4 > data.len() {
        return None;
    }

    if magic == UNCOMP_MAGIC {
        // Uncompressed: just return the raw data after the header
        return Some(data[16..16 + raw_size.min(data.len() - 16)].to_vec());
    }

    if magic != COMP_MAGIC {
        return None;
    }

    // LZ decompression
    let mut dict = [0u8; 4096];
    dict[..INIT_DICT_LEN].copy_from_slice(&INIT_DICT[..INIT_DICT_LEN]);
    let mut write_pos: usize = INIT_DICT_LEN;
    let mut output = Vec::with_capacity(raw_size);
    let mut pos = 16usize; // skip header
    let end = comp_size + 4; // comp_size counts from byte 4

    while pos < end && pos < data.len() {
        // Read control byte
        let control = data[pos];
        pos += 1;

        for i in 0..8 {
            if pos >= end || pos >= data.len() || output.len() >= raw_size {
                break;
            }
            if (control >> i) & 1 == 1 {
                // Dictionary reference (2 bytes)
                if pos + 1 >= data.len() {
                    break;
                }
                let hi = data[pos] as u16;
                let lo = data[pos + 1] as u16;
                pos += 2;

                let offset = ((hi << 4) | (lo >> 4)) as usize;
                let length = (lo & 0x0F) as usize + 2;

                for j in 0..length {
                    if output.len() >= raw_size {
                        break;
                    }
                    let byte = dict[(offset + j) % 4096];
                    output.push(byte);
                    dict[write_pos % 4096] = byte;
                    write_pos += 1;
                }
            } else {
                // Literal byte
                let byte = data[pos];
                pos += 1;
                output.push(byte);
                dict[write_pos % 4096] = byte;
                write_pos += 1;
            }
        }
    }

    Some(output)
}

/// Extract HTML content from decompressed RTF that contains `\fromhtml1`.
///
/// Many `.msg` files store HTML inside RTF using the `\fromhtml1` control word.
/// This function extracts the original HTML by collecting content from
/// `\htmlrtf0` regions and `{\*\htmltag ...}` groups.
///
/// Returns `None` if the RTF does not contain embedded HTML.
pub(crate) fn extract_html_from_rtf(rtf: &[u8]) -> Option<String> {
    let rtf_str = std::str::from_utf8(rtf).ok()?;

    // Check for the \fromhtml1 marker
    if !rtf_str.contains("\\fromhtml1") {
        return None;
    }

    let mut html = String::with_capacity(rtf_str.len());
    let bytes = rtf_str.as_bytes();
    let len = bytes.len();
    let mut i = 0;
    let mut in_htmlrtf = false; // true when inside \htmlrtf (RTF-only region to skip)

    while i < len {
        // Check for \htmlrtf control word
        if bytes[i] == b'\\' && i + 8 < len && &bytes[i..i + 8] == b"\\htmlrtf" {
            i += 8;
            // Check for \htmlrtf0 (end of RTF-only region) vs \htmlrtf (start)
            if i < len && bytes[i] == b'0' {
                in_htmlrtf = false;
                i += 1;
            } else {
                in_htmlrtf = true;
            }
            // Skip optional trailing space
            if i < len && bytes[i] == b' ' {
                i += 1;
            }
            continue;
        }

        // Check for {\*\htmltag ... } groups — these contain actual HTML
        if i + 12 < len && &bytes[i..i + 12] == b"{\\*\\htmltag " {
            i += 12;
            // Skip optional digits (tag number)
            while i < len && bytes[i].is_ascii_digit() {
                i += 1;
            }
            // Skip space after digits
            if i < len && bytes[i] == b' ' {
                i += 1;
            }
            // Collect content until closing '}'
            let mut depth = 1;
            while i < len && depth > 0 {
                match bytes[i] {
                    b'{' => depth += 1,
                    b'}' => {
                        depth -= 1;
                        if depth == 0 {
                            i += 1;
                            break;
                        }
                    }
                    b'\\' if i + 1 < len => {
                        // Handle escaped characters
                        match bytes[i + 1] {
                            b'\\' => {
                                html.push('\\');
                                i += 2;
                                continue;
                            }
                            b'{' => {
                                html.push('{');
                                i += 2;
                                continue;
                            }
                            b'}' => {
                                html.push('}');
                                i += 2;
                                continue;
                            }
                            b'\'' if i + 3 < len => {
                                // \'XX hex escape
                                let hex_str =
                                    std::str::from_utf8(&bytes[i + 2..i + 4]).unwrap_or("3f");
                                if let Ok(byte_val) = u8::from_str_radix(hex_str, 16) {
                                    html.push(byte_val as char);
                                }
                                i += 4;
                                continue;
                            }
                            _ => {
                                // Skip other control words inside htmltag
                                i += 1;
                                while i < len && bytes[i].is_ascii_alphabetic() {
                                    i += 1;
                                }
                                // Skip optional numeric argument
                                if i < len && (bytes[i] == b'-' || bytes[i].is_ascii_digit()) {
                                    i += 1;
                                    while i < len && bytes[i].is_ascii_digit() {
                                        i += 1;
                                    }
                                }
                                // Skip delimiter space
                                if i < len && bytes[i] == b' ' {
                                    i += 1;
                                }
                                continue;
                            }
                        }
                    }
                    _ => {
                        html.push(bytes[i] as char);
                    }
                }
                i += 1;
            }
            continue;
        }

        // Skip content inside \htmlrtf regions (RTF rendering only, not HTML)
        if in_htmlrtf {
            i += 1;
            continue;
        }

        // Skip RTF groups and control words outside of htmltag
        if bytes[i] == b'{' || bytes[i] == b'}' {
            i += 1;
            continue;
        }

        if bytes[i] == b'\\' {
            i += 1;
            // Handle RTF escapes that produce literal characters
            if i < len {
                match bytes[i] {
                    b'\\' => {
                        html.push('\\');
                        i += 1;
                        continue;
                    }
                    b'{' => {
                        html.push('{');
                        i += 1;
                        continue;
                    }
                    b'}' => {
                        html.push('}');
                        i += 1;
                        continue;
                    }
                    b'\'' if i + 2 < len => {
                        let hex_str = std::str::from_utf8(&bytes[i + 1..i + 3]).unwrap_or("3f");
                        if let Ok(byte_val) = u8::from_str_radix(hex_str, 16) {
                            html.push(byte_val as char);
                        }
                        i += 3;
                        continue;
                    }
                    b'\r' | b'\n' => {
                        i += 1;
                        continue;
                    }
                    _ => {
                        // Skip control word
                        while i < len && bytes[i].is_ascii_alphabetic() {
                            i += 1;
                        }
                        // Skip optional numeric parameter
                        if i < len && (bytes[i] == b'-' || bytes[i].is_ascii_digit()) {
                            i += 1;
                            while i < len && bytes[i].is_ascii_digit() {
                                i += 1;
                            }
                        }
                        // Skip delimiter space
                        if i < len && bytes[i] == b' ' {
                            i += 1;
                        }
                        continue;
                    }
                }
            }
            continue;
        }

        // Skip CR/LF (RTF line wrapping, not part of content)
        if bytes[i] == b'\r' || bytes[i] == b'\n' {
            i += 1;
            continue;
        }

        // Plain text outside of control words — part of the HTML
        html.push(bytes[i] as char);
        i += 1;
    }

    if html.is_empty() { None } else { Some(html) }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_decompress_too_short() {
        assert!(decompress_rtf(&[0u8; 10]).is_none());
    }

    #[test]
    fn test_decompress_bad_magic() {
        let mut data = vec![0u8; 20];
        // comp_size = 4 (just header), raw_size = 0, bad magic
        data[0] = 4;
        data[8] = 0xFF;
        assert!(decompress_rtf(&data).is_none());
    }

    #[test]
    fn test_decompress_uncompressed() {
        let content = b"hello world";
        let mut data = Vec::new();
        // comp_size (includes 12 bytes of header after first 4 + content)
        let comp_size = 12 + content.len();
        data.extend(&(comp_size as u32).to_le_bytes());
        data.extend(&(content.len() as u32).to_le_bytes());
        data.extend(&UNCOMP_MAGIC.to_le_bytes());
        data.extend(&0u32.to_le_bytes()); // CRC
        data.extend(content);
        let result = decompress_rtf(&data).unwrap();
        assert_eq!(&result, content);
    }

    #[test]
    fn test_decompress_real_rtf() {
        // Test with actual .msg file RTF
        let outlook = crate::Outlook::from_path("data/test_email.msg").unwrap();
        if !outlook.rtf_compressed.is_empty() {
            let raw = hex::decode(&outlook.rtf_compressed).unwrap();
            let decompressed = decompress_rtf(&raw);
            assert!(decompressed.is_some());
            let rtf = decompressed.unwrap();
            // Decompressed RTF should start with {\rtf
            assert!(
                rtf.starts_with(b"{\\rtf"),
                "RTF should start with {{\\rtf, got: {:?}",
                &rtf[..20.min(rtf.len())]
            );
        }
    }

    #[test]
    fn test_extract_html_no_fromhtml() {
        let rtf = b"{\\rtf1 hello world}";
        assert!(extract_html_from_rtf(rtf).is_none());
    }
}
