#![allow(unused_imports, bare_trait_objects, dead_code)]
//! A simple parser and reader for Microsoft Compound Document File.
//!
//! This includes a basic parser, which validates the format of a given file
//! or a given stream.
//! It includes a reader too, for iterating over entries and for extracting
//! files inside the OLE storage.
//!
//! ## Example
//!
//! ```ignore
//! use crate::ole::Reader;
//! use std::io::{Read, Write};
//!
//! let mut file = std::fs::File::open("data/Thumbs.db").unwrap();
//! let mut parser = Reader::new(file).unwrap();
//!
//! // Iterate through the entries
//! for entry in parser.iterate() {
//!     println!("{}", entry);
//! }
//!
//! // We're going to extract a file from the OLE storage
//! let entry = parser.iterate().next().unwrap();
//! let mut slice = parser.get_entry_slice(entry).unwrap();
//! let mut buffer = std::vec::Vec::<u8>::with_capacity(slice.len());
//! slice.read_to_end(&mut buffer);
//!
//! // Saves the extracted file
//! let mut extracted_file = std::fs::File::create("./file.bin").unwrap();
//! extracted_file.write_all(&buffer[..]);
//! ```

mod ole;
pub use ole::Reader;

pub(crate) mod iterator;
pub(crate) use iterator::OLEIterator;

mod error;
pub use error::Error;

pub(crate) mod header;
pub(crate) mod util;
pub(crate) mod sat;
pub(crate) mod constants;

pub mod entry;
pub use entry::Entry;
pub use entry::EntrySlice;
pub use entry::EntryType;

pub(crate) mod sector;
