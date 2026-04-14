//! # msg_parser
//!
//! A parser for Microsoft Outlook `.msg` files (OLE Compound Document format).
//!
//! Extracts message metadata, body content, recipients, attachments, and
//! transport headers from `.msg` files as specified in
//! [MS-OXMSG](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxmsg)
//! and [MS-OXPROPS](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops).
//!
//! # Quick Start
//!
//! ```no_run
//! use msg_parser::Outlook;
//!
//! let outlook = Outlook::from_path("email.msg").unwrap();
//!
//! println!("From: {} <{}>", outlook.sender.name, outlook.sender.email);
//! println!("Subject: {}", outlook.subject);
//! println!("Date: {}", outlook.message_delivery_time);
//!
//! for person in &outlook.to {
//!     println!("To: {} <{}>", person.name, person.email);
//! }
//!
//! for attach in &outlook.attachments {
//!     println!("Attachment: {} ({} bytes)", attach.long_file_name, attach.payload_bytes.len());
//! }
//! ```
//!
//! # Parsing from different sources
//!
//! ```no_run
//! use msg_parser::Outlook;
//!
//! // From a file path
//! let outlook = Outlook::from_path("email.msg").unwrap();
//!
//! // From a byte slice
//! let bytes = std::fs::read("email.msg").unwrap();
//! let outlook = Outlook::from_slice(&bytes).unwrap();
//!
//! // From any reader (file, stdin, network stream, etc.)
//! let file = std::fs::File::open("email.msg").unwrap();
//! let outlook = Outlook::from_reader(file).unwrap();
//! ```

// OLE Reader
mod ole;

// Outlook Email Message File Parser
mod parser;
pub use parser::*;
