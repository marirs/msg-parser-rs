use std::io;

use serde_json::Error as SerdeError;

use thiserror::Error as ThisError;

use crate::ole::Error as OleError;

// DataTypeError is used when decode fails in datatype.rs
#[derive(ThisError, Debug)]
pub enum DataTypeError {
    #[error("DataTypeError: Unknown value encoding: 0x{0}")]
    UnknownCode(String),
    #[error("DataTypeError: Unable to decode bytes into UTF-8 string {0}")]
    Utf8Err(#[from] std::string::FromUtf8Error),
    #[error("DataTypeError: Unable to decode bytes into UTF-16 string {0}")]
    Utf16Err(#[from] std::string::FromUtf16Error),
}

#[derive(ThisError, Debug)]
pub enum Error {
    #[error(transparent)]
    DataTypeError(#[from] DataTypeError),

    #[error("Unable to read file")]
    Io {
        #[from]
        source: io::Error,
    },

    #[error("Error parsing file with ole: {}", .source)]
    OleError {
        #[from]
        source: OleError,
    },

    #[error(transparent)]
    SerdeJsonError(#[from] SerdeError),
}
