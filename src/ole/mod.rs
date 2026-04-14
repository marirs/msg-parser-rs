#![allow(unused_imports, dead_code)]
//! A simple parser and reader for Microsoft Compound Document File.
//!
//! This includes a basic parser, which validates the format of a given file
//! or a given stream.
//! It includes a reader too, for iterating over entries and for extracting
//! files inside the OLE storage.
//!
//! ## Example
//!
//! ```rust,ignore
//! use crate::ole::Reader;
//! use std::io::{Read, Write};
//!
//! let parser = Reader::from_path("data/test_email.msg").unwrap();
//!
//! // Iterate through the entries
//! for entry in parser.iterate() {
//!     println!("{}", entry);
//! }
//!
//! // Extract a stream from the OLE storage
//! let entry = parser.iterate().next().unwrap();
//! let mut slice = parser.get_entry_slice(entry).unwrap();
//! let mut buffer = Vec::<u8>::with_capacity(slice.len());
//! slice.read_to_end(&mut buffer).unwrap();
//!
//! // Save the extracted stream
//! let mut extracted_file = std::fs::File::create("./file.bin").unwrap();
//! extracted_file.write_all(&buffer).unwrap();
//! ```

#[allow(clippy::module_inception)]
mod ole;
pub use ole::Reader;

pub(crate) mod iterator;
pub(crate) use iterator::OLEIterator;

mod error;
pub use error::Error;

pub(crate) mod constants;
pub(crate) mod header;
pub(crate) mod sat;
pub(crate) mod util;

pub mod entry;
pub use entry::Entry;
pub use entry::EntrySlice;
pub use entry::EntryType;

pub(crate) mod sector;
