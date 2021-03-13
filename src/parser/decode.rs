use std::io::Read;

use hex;

use crate::ole::EntrySlice;

use super::error::{DataTypeError, Error};

// DataType corresponds to decoded property values
// as specified in this document.
// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/0c77892e-288e-435a-9c49-be1c20c7afdb
#[derive(Clone, Debug, PartialEq)]
pub enum DataType {
    PtypString(String),
    PtypBinary(Vec<u8>),
}

impl From<&DataType> for String {
    fn from(data: &DataType) -> Self {
        match *data {
            DataType::PtypBinary(ref bytes) => hex::encode(bytes),
            DataType::PtypString(ref string) => string.to_string(),
        }
    }
}

// PytpDecoder converts a byte sequence
// into primitive type DataType.
pub struct PtypDecoder {}

impl PtypDecoder {
    pub fn decode(entry_slice: &mut EntrySlice, code: &str) -> Result<DataType, Error> {
        let mut buff = vec![0u8; entry_slice.len()];
        entry_slice.read(&mut buff)?;
        match code {
            "0x001F" => decode_ptypstring(&buff),
            "0x0102" => decode_ptypbinary(&buff),
            _ => Err(DataTypeError::UnknownCode(code.to_string()).into()),
        }
    }
}

fn decode_ptypbinary(buff: &Vec<u8>) -> Result<DataType, Error> {
    Ok(DataType::PtypBinary(buff.to_vec()))
}

fn decode_ptypstring(buff: &Vec<u8>) -> Result<DataType, Error> {
    // PtypString
    // Byte sequence is in little-endian format
    // Use UTF-16 String decode
    let buffu16: Vec<u16> = if cfg!(target_endian = "little") {
        buff.iter().map(|&x| x as u16).collect()
    } else {
        buff.iter().map(|&x| x as u16).rev().collect()
    };
    match String::from_utf16(&buffu16) {
        // Remove all terminated null character
        Ok(decoded) => Ok(DataType::PtypString(decoded.replace(char::from(0), ""))),
        Err(err) => Err(DataTypeError::Utf16Err(err).into()),
    }
}

#[cfg(test)]
mod tests {
    use super::{DataType, PtypDecoder};
    use crate::ole::Reader;

    #[test]
    fn test_unknown_code() {
        // Test with dummy file.
        let path = "data/test_email.msg";
        let parser = Reader::from_path(path).unwrap();
        let entry = parser.iterate().next().unwrap();

        let mut slice = parser.get_entry_slice(entry).unwrap();
        let res = PtypDecoder::decode(&mut slice, "1234");
        assert_eq!(res.is_err(), true);
        let err = res.unwrap_err();
        assert_eq!(
            err.to_string(),
            "DataTypeError: Unknown value encoding: 0x1234"
        );
    }

    #[test]
    fn test_ptypstring() {
        let path = "data/test_email.msg";
        let parser = Reader::from_path(path).unwrap();

        let entry_of_a_ptypstring = parser.iterate().nth(125).unwrap();
        let mut ptypstring_slice = parser.get_entry_slice(entry_of_a_ptypstring).unwrap();
        let ptypstring_decoded = PtypDecoder::decode(&mut ptypstring_slice, "0x001F").unwrap();
        assert_eq!(
            ptypstring_decoded,
            DataType::PtypString("marirs@outlook.com".to_string())
        );
    }
}
