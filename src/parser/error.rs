use std::{
    io,
};

use serde_json::Error as SerdeError;

use thiserror::Error as ThisError;

use crate::ole::Error as OleError;

// DataTypeError is used when decode fails in datatype.rs
#[derive(ThisError, Debug)]
pub enum DataTypeError {
    UnknownCode(String),
    Utf8Err(#[from] std::string::FromUtf8Error),
    Utf16Err(#[from] std::string::FromUtf16Error),
}

impl std::fmt::Display for DataTypeError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match *self {
            DataTypeError::UnknownCode(ref value) => {
                write!(f, "DataTypeError: Unknown value encoding: 0x{}", value)
            }
            DataTypeError::Utf8Err(ref err) => {
                write!(
                    f,
                    "DataTypeError: Unable to decode bytes into UTF-8 string {}",
                    err.to_string()
                )
            }
            DataTypeError::Utf16Err(ref err) => {
                write!(
                    f,
                    "DataTypeError: Unable to decode bytes into UTF-16 string {}",
                    err.to_string()
                )
            }
        }
    }
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
