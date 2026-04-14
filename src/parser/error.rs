use std::io;

use serde_json::Error as SerdeError;

use thiserror::Error as ThisError;

use crate::ole::Error as OleError;

/// Error type for property value decoding failures.
#[derive(ThisError, Debug)]
pub enum DataTypeError {
    /// The property type code is not recognized.
    #[error("DataTypeError: Unknown value encoding: 0x{0}")]
    UnknownCode(String),
    /// Failed to decode bytes as a UTF-8 string.
    #[error("DataTypeError: Unable to decode bytes into UTF-8 string {0}")]
    Utf8Err(#[from] std::string::FromUtf8Error),
    /// Failed to decode bytes as a UTF-16 string.
    #[error("DataTypeError: Unable to decode bytes into UTF-16 string {0}")]
    Utf16Err(#[from] std::string::FromUtf16Error),
}

/// Top-level error type returned by [`Outlook`](crate::Outlook) parsing methods.
#[derive(ThisError, Debug)]
pub enum Error {
    /// A property value could not be decoded.
    #[error(transparent)]
    DataTypeError(#[from] DataTypeError),

    /// An I/O error occurred while reading the file or stream.
    #[error("Unable to read file")]
    Io {
        #[from]
        source: io::Error,
    },

    /// The underlying OLE compound document could not be parsed.
    #[error("Error parsing file with ole: {}", .source)]
    OleError {
        #[from]
        source: OleError,
    },

    /// JSON serialization failed (from [`Outlook::to_json`](crate::Outlook::to_json)).
    #[error(transparent)]
    SerdeJsonError(#[from] SerdeError),
}
