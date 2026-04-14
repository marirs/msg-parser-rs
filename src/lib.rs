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
//! // Use Display impl for a human-readable summary
//! println!("{}", outlook);
//!
//! // Or access fields directly
//! println!("From: {}", outlook.sender);
//! println!("Subject: {}", outlook.subject);
//! println!("Date: {}", outlook.message_delivery_time);
//!
//! for attach in &outlook.attachments {
//!     println!("Attachment: {}", attach);
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
//! // From a byte slice (accepts &[u8], Vec<u8>, or anything AsRef<[u8]>)
//! let bytes = std::fs::read("email.msg").unwrap();
//! let outlook = Outlook::from_slice(&bytes).unwrap();
//!
//! // From any reader (file, stdin, network stream, etc.)
//! let file = std::fs::File::open("email.msg").unwrap();
//! let outlook = Outlook::from_reader(file).unwrap();
//! ```
//!
//! # Embedded messages
//!
//! ```no_run
//! # let outlook = msg_parser::Outlook::from_path("email.msg").unwrap();
//! for attach in &outlook.attachments {
//!     if let Some(Ok(nested)) = attach.as_message() {
//!         println!("Embedded: {} from {}", nested.subject, nested.sender);
//!     }
//! }
//! ```
//!
//! # HTML from RTF fallback
//!
//! ```no_run
//! # let outlook = msg_parser::Outlook::from_path("email.msg").unwrap();
//! let html = if !outlook.html.is_empty() {
//!     outlook.html.clone()
//! } else {
//!     outlook.html_from_rtf().unwrap_or_default()
//! };
//! ```

// OLE Reader
mod ole;

// Outlook Email Message File Parser
mod parser;
pub use parser::*;
