use std::{io::Read, path::Path, sync::LazyLock};

use regex::Regex;

use serde::{Deserialize, Serialize};

use crate::ole;

use super::{
    error::Error,
    storage::{Properties, Storages},
};

type Name = String;
type Email = String;

static RE_CONTENT_TYPE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?im)^Content-Type: (.*(\n\s.*)*)\r\n").unwrap());
static RE_DATE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?i)Date: (.*(\n\s.*)*)\r\n").unwrap());
static RE_MESSAGE_ID: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?im)^Message-ID: (.*(\n\s.*)*)\r\n").unwrap());
static RE_REPLY_TO: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(?im)^Reply-To: (.*(\n\s.*)*)\r\n").unwrap());

/// SMTP transport headers from the message envelope.
///
/// Contains a few commonly-used headers extracted via regex, plus the full
/// raw header text for custom parsing. Messages that were never sent via
/// SMTP (e.g. drafts) will have all fields empty.
#[non_exhaustive]
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct TransportHeaders {
    /// The full, unparsed transport header text.
    pub raw: String,
    /// The `Content-Type` header value.
    pub content_type: String,
    /// The `Date` header value.
    pub date: String,
    /// The `Message-ID` header value.
    pub message_id: String,
    /// The `Reply-To` header value.
    pub reply_to: String,
}

impl TransportHeaders {
    fn extract_field(text: &str, re: &Regex) -> String {
        if text.is_empty() {
            return String::new();
        }
        re.captures(text)
            .and_then(|cap| cap.get(1).map(|x| String::from(x.as_str())))
            .unwrap_or_default()
    }

    /// Parse transport headers from raw header text.
    pub fn create_from_headers_text(text: &str) -> Self {
        Self {
            raw: text.to_string(),
            content_type: Self::extract_field(text, &RE_CONTENT_TYPE),
            date: Self::extract_field(text, &RE_DATE),
            message_id: Self::extract_field(text, &RE_MESSAGE_ID),
            reply_to: Self::extract_field(text, &RE_REPLY_TO),
        }
    }
}

/// A person referenced in the message (sender, recipient, CC, or BCC).
#[non_exhaustive]
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct Person {
    /// Display name (e.g. `"John Doe"`).
    pub name: Name,
    /// SMTP email address (e.g. `"john@example.com"`).
    pub email: Email,
}

impl Person {
    #[cfg(test)]
    fn new(name: Name, email: Email) -> Self {
        Self { name, email }
    }
    fn create_from_props(props: &Properties, name_key: &str, email_keys: &[&str]) -> Self {
        let name: String = props.get(name_key).map_or(String::new(), |x| x.into());
        // Get the first email that can be found in props given email_keys.
        let email = email_keys
            .iter()
            .map(|&key| props.get(key).map_or(String::new(), |x| x.into()))
            .find(|x| !x.is_empty())
            .unwrap_or_default();
        Self { name, email }
    }

    /// Returns `true` if the email address looks like an Exchange X.500 DN
    /// rather than a proper SMTP address.
    fn is_x500_dn(email: &str) -> bool {
        let upper = email.to_ascii_uppercase();
        upper.starts_with("/O=") || upper.starts_with("/CN=")
    }

    /// Try to resolve X.500 DN addresses to SMTP by searching the raw
    /// transport headers for the display name.
    fn resolve_email(&mut self, raw_headers: &str) {
        if self.email.is_empty() || !Self::is_x500_dn(&self.email) {
            return;
        }
        // Try to find this person's SMTP address in the transport headers.
        // Headers often contain: "Display Name" <user@example.com>
        // Look for the display name (or the /CN= tail) followed by an SMTP address.
        if let Some(smtp) = Self::find_smtp_in_headers(raw_headers, &self.name) {
            self.email = smtp;
        }
    }

    /// Search raw transport headers for an SMTP address associated with a display name.
    fn find_smtp_in_headers(headers: &str, display_name: &str) -> Option<String> {
        if display_name.is_empty() || headers.is_empty() {
            return None;
        }
        // Look for patterns like: "Display Name" <user@domain.com>
        // or: Display Name <user@domain.com>
        let name_lower = display_name.to_lowercase();
        for line in headers.lines() {
            let line_lower = line.to_lowercase();
            if !line_lower.contains(&name_lower) {
                continue;
            }
            // Extract email from angle brackets
            if let Some(start) = line.rfind('<')
                && let Some(end) = line[start..].find('>')
            {
                let candidate = &line[start + 1..start + end];
                if candidate.contains('@') {
                    return Some(candidate.to_string());
                }
            }
        }
        None
    }
}

/// A file attachment on the message.
///
/// # Saving an attachment to disk
///
/// ```no_run
/// # let outlook = msg_parser::Outlook::from_path("email.msg").unwrap();
/// for attach in &outlook.attachments {
///     let name = if attach.long_file_name.is_empty() {
///         &attach.file_name
///     } else {
///         &attach.long_file_name
///     };
///     std::fs::write(name, &attach.payload_bytes).unwrap();
/// }
/// ```
#[non_exhaustive]
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct Attachment {
    /// Display name shown in the mail client.
    pub display_name: String,
    /// Attachment content as a hex-encoded string. Use [`payload_bytes`](Attachment::payload_bytes) for raw bytes.
    pub payload: String,
    /// Attachment content as raw bytes. Identical data to `payload`, just not hex-encoded.
    #[serde(with = "hex")]
    pub payload_bytes: Vec<u8>,
    /// File extension including the dot (e.g. `".pdf"`).
    pub extension: String,
    /// MIME type (e.g. `"image/png"`). May be empty.
    pub mime_tag: String,
    /// Short 8.3 filename (e.g. `"docume~1.pdf"`).
    pub file_name: String,
    /// Full original filename (e.g. `"document_final.pdf"`).
    pub long_file_name: String,
    /// How the attachment is stored. Common values:
    /// - `1` — by value (regular file, bytes in `payload_bytes`)
    /// - `5` — embedded message (nested `.msg`)
    /// - `6` — OLE object
    pub attach_method: u32,
    /// Content-ID for inline attachments (e.g. `"image001@01D00000.00000000"`).
    /// Used to resolve `cid:` references in HTML bodies. Empty if not set.
    pub content_id: String,
}

impl Attachment {
    fn create(storages: &Storages, idx: usize) -> Self {
        let payload_bytes = storages.get_bytes_from_attachment(idx, "AttachDataObject");
        let payload = hex::encode(&payload_bytes);
        Self {
            display_name: storages.get_val_from_attachment_or_default(idx, "DisplayName"),
            payload,
            payload_bytes,
            extension: storages.get_val_from_attachment_or_default(idx, "AttachExtension"),
            mime_tag: storages.get_val_from_attachment_or_default(idx, "AttachMimeTag"),
            file_name: storages.get_val_from_attachment_or_default(idx, "AttachFilename"),
            long_file_name: storages.get_val_from_attachment_or_default(idx, "AttachLongFilename"),
            attach_method: storages
                .get_attachment_int_prop(idx, "AttachMethod")
                .unwrap_or(0),
            content_id: storages.get_val_from_attachment_or_default(idx, "AttachContentId"),
        }
    }

    /// Returns `true` if this attachment is an embedded `.msg` message
    /// (`attach_method == 5`).
    pub fn is_embedded_message(&self) -> bool {
        self.attach_method == 5
    }

    /// Parse the embedded `.msg` attachment and return the nested message.
    ///
    /// Returns `Some(Ok(outlook))` if this is an embedded message (`attach_method == 5`)
    /// with parseable content, `Some(Err(_))` if it is an embedded message but parsing
    /// fails, or `None` if it is not an embedded message.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let outlook = Outlook::from_path("email.msg").unwrap();
    /// for attach in &outlook.attachments {
    ///     if let Some(Ok(nested)) = attach.as_message() {
    ///         println!("Embedded: {}", nested.subject);
    ///     }
    /// }
    /// ```
    pub fn as_message(&self) -> Option<Result<Outlook, Error>> {
        if !self.is_embedded_message() || self.payload_bytes.is_empty() {
            return None;
        }
        Some(Outlook::from_slice(&self.payload_bytes))
    }
}

/// A parsed Outlook `.msg` email message.
///
/// Create an instance with [`from_path`](Outlook::from_path),
/// [`from_slice`](Outlook::from_slice), or [`from_reader`](Outlook::from_reader),
/// then access the message fields directly.
///
/// Field names follow the
/// [MS-OXPROPS](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/f6ab1613-aefe-447d-a49c-18217230b148)
/// specification.
///
/// # Example
///
/// ```no_run
/// use msg_parser::Outlook;
///
/// let outlook = Outlook::from_path("email.msg").unwrap();
/// println!("Subject: {}", outlook.subject);
/// println!("From: {} <{}>", outlook.sender.name, outlook.sender.email);
/// println!("To: {:?}", outlook.to);
/// println!("CC: {:?}", outlook.cc);
/// println!("BCC: {:?}", outlook.bcc);
/// println!("Delivered: {}", outlook.message_delivery_time);
/// println!("Attachments: {}", outlook.attachments.len());
/// ```
#[non_exhaustive]
#[derive(Serialize, Deserialize, Debug)]
pub struct Outlook {
    /// SMTP transport headers. Empty for messages that were never sent.
    pub headers: TransportHeaders,
    /// Message sender.
    pub sender: Person,
    /// Primary recipients (`RecipientType = 1`).
    pub to: Vec<Person>,
    /// Carbon-copy recipients (`RecipientType = 2`).
    pub cc: Vec<Person>,
    /// Blind carbon-copy recipients (`RecipientType = 3`).
    pub bcc: Vec<Person>,
    /// Message subject line.
    pub subject: String,
    /// Plain-text body.
    pub body: String,
    /// HTML body. Empty if the message has no HTML representation.
    pub html: String,
    /// RTF compressed body (hex-encoded).
    pub rtf_compressed: String,
    /// Message class, typically `"IPM.Note"` for regular emails.
    pub message_class: String,
    /// Importance level: `0` = Low, `1` = Normal, `2` = High.
    pub importance: u32,
    /// Sensitivity level: `0` = Normal, `1` = Personal, `2` = Private, `3` = Confidential.
    pub sensitivity: u32,
    /// When the sender submitted the message (ISO 8601 UTC). Empty if unavailable.
    pub client_submit_time: String,
    /// When the message was delivered (ISO 8601 UTC). Empty if unavailable.
    pub message_delivery_time: String,
    /// When the message object was created (ISO 8601 UTC). Empty if unavailable.
    pub creation_time: String,
    /// When the message was last modified (ISO 8601 UTC). Empty if unavailable.
    pub last_modification_time: String,
    /// File attachments. See [`Attachment`] for details.
    pub attachments: Vec<Attachment>,
}

impl Outlook {
    fn populate(storages: &Storages) -> Self {
        let headers_text = storages.get_val_from_root_or_default("TransportMessageHeaders");
        let headers = TransportHeaders::create_from_headers_text(&headers_text);

        let mut to = Vec::new();
        let mut cc = Vec::new();
        let mut bcc = Vec::new();

        for (i, recip_map) in storages.recipients.iter().enumerate() {
            let mut person = Person::create_from_props(
                recip_map,
                "DisplayName",
                &["SmtpAddress", "EmailAddress"],
            );
            // Resolve X.500 DN addresses to SMTP via transport headers
            person.resolve_email(&headers_text);
            // RecipientType: 1=To, 2=CC, 3=BCC (MS-OXMSG 2.2.1)
            match storages.get_recipient_int_prop(i, "RecipientType") {
                Some(2) => cc.push(person),
                Some(3) => bcc.push(person),
                _ => to.push(person), // Default to To (including type==1 and missing)
            }
        }

        let mut sender = Person::create_from_props(
            &storages.root,
            "SenderName",
            &["SenderSmtpAddress", "SenderEmailAddress"],
        );
        sender.resolve_email(&headers_text);

        Self {
            headers,
            sender,
            to,
            cc,
            bcc,
            subject: storages.get_val_from_root_or_default("Subject"),
            body: storages.get_val_from_root_or_default("Body"),
            html: storages.get_val_from_root_or_default("Html"),
            rtf_compressed: storages.get_val_from_root_or_default("RtfCompressed"),
            message_class: storages.get_val_from_root_or_default("MessageClass"),
            importance: storages.get_root_int_prop("Importance").unwrap_or(1),
            sensitivity: storages.get_root_int_prop("Sensitivity").unwrap_or(0),
            client_submit_time: storages.get_val_from_root_or_default("ClientSubmitTime"),
            message_delivery_time: storages.get_val_from_root_or_default("MessageDeliveryTime"),
            creation_time: storages.get_val_from_root_or_default("CreationTime"),
            last_modification_time: storages.get_val_from_root_or_default("LastModificationTime"),
            attachments: storages
                .attachments
                .iter()
                .enumerate()
                .map(|(i, _)| Attachment::create(storages, i))
                .collect(),
        }
    }

    /// Parse a `.msg` file from a filesystem path.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let outlook = Outlook::from_path("email.msg").unwrap();
    /// println!("{}", outlook.subject);
    /// ```
    pub fn from_path<P: AsRef<Path>>(path: P) -> Result<Self, Error> {
        let data = std::fs::read(path)?;
        let parser = ole::Reader::from_bytes(data)?;
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);

        let outlook = Self::populate(&storages);
        Ok(outlook)
    }

    /// Parse a `.msg` file from any [`Read`](std::io::Read) source.
    ///
    /// Reads the entire stream into memory, then parses. Useful for stdin,
    /// network streams, or any non-seekable source.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let file = std::fs::File::open("email.msg").unwrap();
    /// let outlook = Outlook::from_reader(file).unwrap();
    /// ```
    pub fn from_reader<R: std::io::Read>(reader: R) -> Result<Self, Error> {
        // Cap reads at 256 MB to prevent unbounded memory allocation
        const MAX_SIZE: u64 = 256 * 1024 * 1024;
        let mut limited = reader.take(MAX_SIZE + 1);
        let mut buf = Vec::new();
        limited.read_to_end(&mut buf)?;
        if buf.len() as u64 > MAX_SIZE {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "Input exceeds maximum allowed size (256 MB)",
            )
            .into());
        }
        // Use from_bytes directly to avoid an extra copy through from_slice
        let parser = ole::Reader::from_bytes(buf)?;
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);
        Ok(Self::populate(&storages))
    }

    /// Parse a `.msg` file from a byte slice already in memory.
    ///
    /// Accepts any type that implements `AsRef<[u8]>`, including `&[u8]`,
    /// `Vec<u8>`, and `bytes::Bytes`.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let bytes = std::fs::read("email.msg").unwrap();
    /// let outlook = Outlook::from_slice(&bytes).unwrap();
    /// ```
    pub fn from_slice(slice: impl AsRef<[u8]>) -> Result<Self, Error> {
        let parser = ole::Reader::from_bytes(slice.as_ref())?;
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);

        let outlook = Self::populate(&storages);
        Ok(outlook)
    }

    /// Serialize the parsed message to a JSON string.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let outlook = Outlook::from_path("email.msg").unwrap();
    /// let json = outlook.to_json().unwrap();
    /// println!("{}", json);
    /// ```
    pub fn to_json(&self) -> Result<String, Error> {
        Ok(serde_json::to_string(self)?)
    }

    /// Decompress the RTF body and return it as a byte vector.
    ///
    /// The `rtf_compressed` field contains the raw compressed data as a
    /// hex-encoded string. This method decodes and decompresses it per
    /// [MS-OXRTFCP](https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp).
    ///
    /// Returns `None` if the message has no RTF body or decompression fails.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let outlook = Outlook::from_path("email.msg").unwrap();
    /// if let Some(rtf) = outlook.rtf_decompressed() {
    ///     println!("RTF body: {} bytes", rtf.len());
    /// }
    /// ```
    pub fn rtf_decompressed(&self) -> Option<Vec<u8>> {
        if self.rtf_compressed.is_empty() {
            return None;
        }
        let raw = hex::decode(&self.rtf_compressed).ok()?;
        super::rtf::decompress_rtf(&raw)
    }

    /// Extract HTML from the RTF body when the message has no direct HTML property.
    ///
    /// Many Outlook messages embed the HTML body inside compressed RTF using
    /// the `\fromhtml1` control word. This method decompresses the RTF and
    /// extracts the embedded HTML.
    ///
    /// Returns `None` if the RTF body doesn't contain embedded HTML.
    /// Prefer the `html` field when it is non-empty — this method is a fallback.
    ///
    /// # Example
    ///
    /// ```no_run
    /// use msg_parser::Outlook;
    ///
    /// let outlook = Outlook::from_path("email.msg").unwrap();
    /// let html = if !outlook.html.is_empty() {
    ///     outlook.html.clone()
    /// } else {
    ///     outlook.html_from_rtf().unwrap_or_default()
    /// };
    /// ```
    pub fn html_from_rtf(&self) -> Option<String> {
        let rtf = self.rtf_decompressed()?;
        super::rtf::extract_html_from_rtf(&rtf)
    }
}

impl std::fmt::Display for Person {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        if self.name.is_empty() {
            write!(f, "{}", self.email)
        } else if self.email.is_empty() || self.name == self.email {
            write!(f, "{}", self.name)
        } else {
            write!(f, "{} <{}>", self.name, self.email)
        }
    }
}

impl std::fmt::Display for Attachment {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = if self.long_file_name.is_empty() {
            if self.file_name.is_empty() {
                &self.display_name
            } else {
                &self.file_name
            }
        } else {
            &self.long_file_name
        };
        let method = match self.attach_method {
            1 => "file",
            5 => "embedded .msg",
            6 => "OLE object",
            _ => "unknown",
        };
        write!(
            f,
            "{} ({}, {} bytes)",
            name,
            method,
            self.payload_bytes.len()
        )
    }
}

impl std::fmt::Display for Outlook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "From:    {}", self.sender)?;
        writeln!(f, "Subject: {}", self.subject)?;
        if !self.to.is_empty() {
            let to: Vec<String> = self.to.iter().map(|p| p.to_string()).collect();
            writeln!(f, "To:      {}", to.join(", "))?;
        }
        if !self.cc.is_empty() {
            let cc: Vec<String> = self.cc.iter().map(|p| p.to_string()).collect();
            writeln!(f, "CC:      {}", cc.join(", "))?;
        }
        if !self.bcc.is_empty() {
            let bcc: Vec<String> = self.bcc.iter().map(|p| p.to_string()).collect();
            writeln!(f, "BCC:     {}", bcc.join(", "))?;
        }
        if !self.message_delivery_time.is_empty() {
            writeln!(f, "Date:    {}", self.message_delivery_time)?;
        }
        if !self.attachments.is_empty() {
            writeln!(f, "Attachments ({}):", self.attachments.len())?;
            for a in &self.attachments {
                writeln!(f, "  - {}", a)?;
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{Outlook, Person, TransportHeaders};

    #[test]
    fn test_invalid_file() {
        let path = "data/bad_outlook.msg";
        let err = Outlook::from_path(path).unwrap_err();
        assert_eq!(
            err.to_string(),
            "Error parsing file with ole: Invalid OLE File".to_string()
        );
    }

    #[test]
    fn test_transport_header_test_email_1() {
        use super::super::storage::Storages;
        use crate::ole::Reader;

        let parser = Reader::from_path("data/test_email.msg").unwrap();
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);

        let transport_text = storages.get_val_from_root_or_default("TransportMessageHeaders");

        let header = TransportHeaders::create_from_headers_text(&transport_text);

        assert!(header.raw.is_empty());
        assert!(header.content_type.is_empty());
        assert!(header.date.is_empty());
        assert!(header.message_id.is_empty());
        assert!(header.reply_to.is_empty());
    }

    #[test]
    fn test_test_email() {
        let path = "data/test_email.msg";
        let outlook = Outlook::from_path(path).unwrap();
        assert_eq!(
            outlook.sender,
            Person {
                name: "".to_string(),
                email: "".to_string()
            }
        );

        // RecipientType == 1 (To)
        assert_eq!(
            outlook.to,
            vec![Person {
                name: "marirs@outlook.com".to_string(),
                email: "marirs@outlook.com".to_string()
            }]
        );

        // RecipientType == 2 (CC)
        assert_eq!(
            outlook.cc,
            vec![
                Person {
                    name: "Sriram Govindan".to_string(),
                    email: "marirs@aol.in".to_string()
                },
                Person {
                    name: "marirs@outlook.in".to_string(),
                    email: "marirs@outlook.in".to_string()
                },
            ]
        );

        // RecipientType == 3 (BCC)
        assert_eq!(
            outlook.bcc,
            vec![
                Person {
                    name: "Sriram Govindan".to_string(),
                    email: "marirs@aol.in".to_string()
                },
                Person {
                    name: "Sriram Govindan".to_string(),
                    email: "marirs@outlook.com".to_string()
                },
                Person {
                    name: "marirs@outlook.in".to_string(),
                    email: "marirs@outlook.in".to_string()
                },
            ]
        );

        assert_eq!(outlook.subject, String::from("Test Email"));

        assert!(outlook.headers.raw.is_empty());
        assert!(outlook.headers.content_type.is_empty());

        assert!(outlook.body.starts_with("Test Email\r\n"));
        assert!(outlook.rtf_compressed.starts_with("51210000c8a200004c5a4"));
    }

    #[test]
    fn test_test_email_2() {
        let path = "data/test_email.msg";
        let outlook = Outlook::from_path(path).unwrap();
        assert_eq!(
            outlook.sender,
            Person {
                name: "".to_string(),
                email: "".to_string()
            }
        );
        assert_eq!(outlook.to.len(), 1);
        assert_eq!(outlook.cc.len(), 2);
        assert_eq!(outlook.bcc.len(), 3);
        assert_eq!(outlook.subject, String::from("Test Email"));

        assert!(outlook.body.starts_with("Test Email"));

        assert_eq!(outlook.attachments.len(), 3);
        // Check displaynames
        let displays: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.display_name.clone())
            .collect();
        assert_eq!(
            displays,
            vec![
                "1 Days Left—35% off cloud space, upgrade now!".to_string(),
                "milky-way-2695569_960_720.jpg".to_string(),
                "Test Email.msg".to_string(),
            ]
        );
        // Check extensions
        let exts: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.extension.clone())
            .collect();
        assert_eq!(
            exts,
            vec!["".to_string(), ".jpg".to_string(), ".msg".to_string()]
        );
        // Check mime tag
        let mimes: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.mime_tag.clone())
            .collect();
        assert_eq!(mimes, vec!["".to_string(), "".to_string(), "".to_string()]);
        // Check filenames
        let filenames: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.file_name.clone())
            .collect();
        assert_eq!(
            filenames,
            vec![
                "".to_string(),
                "milky-~1.jpg".to_string(),
                "TestEm~1.msg".to_string()
            ]
        );
    }

    #[test]
    fn test_attachment_msg() {
        let path = "data/attachment.msg";
        let outlook = Outlook::from_path(path).unwrap();
        assert_eq!(outlook.attachments.len(), 3);

        // Check displaynames
        let displays: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.display_name.clone())
            .collect();
        assert_eq!(
            displays,
            vec![
                "loan_proposal.doc".to_string(),
                "image001.png".to_string(),
                "image002.jpg".to_string()
            ]
        );
        // Check extensions
        let exts: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.extension.clone())
            .collect();
        assert_eq!(
            exts,
            vec![".doc".to_string(), ".png".to_string(), ".jpg".to_string()]
        );
        // Check mime tag
        let mimes: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.mime_tag.clone())
            .collect();
        assert_eq!(
            mimes,
            vec![
                "application/msword".to_string(),
                "image/png".to_string(),
                "image/jpeg".to_string()
            ]
        );
        // Check filenames
        let filenames: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.file_name.clone())
            .collect();
        assert_eq!(
            filenames,
            vec![
                "loan_p~1.doc".to_string(),
                "image001.png".to_string(),
                "image002.jpg".to_string()
            ]
        );
        // Check long filenames
        let long_names: Vec<String> = outlook
            .attachments
            .iter()
            .map(|x| x.long_file_name.clone())
            .collect();
        assert_eq!(
            long_names,
            vec![
                "loan_proposal.doc".to_string(),
                "image001.png".to_string(),
                "image002.jpg".to_string()
            ]
        );
    }

    #[test]
    fn test_payload_bytes() {
        let outlook = Outlook::from_path("data/attachment.msg").unwrap();

        // payload_bytes contains raw binary, payload contains hex encoding
        for attach in &outlook.attachments {
            assert_eq!(attach.payload, hex::encode(&attach.payload_bytes));
        }

        // Verify magic bytes: DOC (OLE), PNG, JPEG
        assert_eq!(
            &outlook.attachments[0].payload_bytes[..4],
            b"\xd0\xcf\x11\xe0"
        );
        assert_eq!(&outlook.attachments[1].payload_bytes[..4], b"\x89PNG");
        assert_eq!(&outlook.attachments[2].payload_bytes[..2], b"\xff\xd8");

        // Non-empty
        assert!(!outlook.attachments[0].payload_bytes.is_empty());
        assert!(!outlook.attachments[1].payload_bytes.is_empty());
        assert!(!outlook.attachments[2].payload_bytes.is_empty());
    }

    #[test]
    fn test_attach_method() {
        // test_email.msg: attachment #0 is embedded msg (method=5), #1 and #2 are by_value (method=1)
        let outlook = Outlook::from_path("data/test_email.msg").unwrap();
        assert_eq!(outlook.attachments[0].attach_method, 5); // embedded_msg
        assert_eq!(outlook.attachments[1].attach_method, 1); // by_value
        assert_eq!(outlook.attachments[2].attach_method, 1); // by_value

        // attachment.msg: all by_value
        let outlook = Outlook::from_path("data/attachment.msg").unwrap();
        for attach in &outlook.attachments {
            assert_eq!(attach.attach_method, 1);
        }
    }

    #[test]
    fn test_unicode_msg() {
        let path = "data/unicode.msg";
        let outlook = Outlook::from_path(path).unwrap();
        assert_eq!(
            outlook.sender,
            Person {
                name: "Brian Zhou".to_string(),
                email: "brizhou@gmail.com".to_string()
            }
        );
        // Recipient #0 is To
        assert_eq!(
            outlook.to,
            vec![Person {
                name: "brianzhou@me.com".to_string(),
                email: "brianzhou@me.com".to_string()
            }]
        );

        // Recipient #1 is CC
        assert_eq!(
            outlook.cc,
            vec![Person::new(
                "Brian Zhou".to_string(),
                "brizhou@gmail.com".to_string()
            )]
        );

        assert!(outlook.bcc.is_empty());
        assert_eq!(outlook.subject, String::from("Test for TIF files"));
        assert!(!outlook.headers.raw.is_empty());
        assert_eq!(
            outlook.headers.content_type,
            "multipart/mixed; boundary=001a113392ecbd7a5404eb6f4d6a"
        );
        assert_eq!(outlook.headers.date, "Mon, 18 Nov 2013 10:26:24 +0200");
        assert_eq!(
            outlook.headers.message_id,
            "<CADtJ4eNjQSkGcBtVteCiTF+YFG89+AcHxK3QZ=-Mt48xygkvdQ@mail.gmail.com>"
        );
        assert!(outlook.headers.reply_to.is_empty());
        assert!(outlook.rtf_compressed.starts_with("bc020000b908"));
    }

    #[test]
    fn test_ascii() {
        let path = "data/ascii.msg";
        let outlook = Outlook::from_path(path).unwrap();
        assert_eq!(
            outlook.sender,
            Person {
                name: "from@domain.com".to_string(),
                email: "from@domain.com".to_string()
            }
        );
        assert_eq!(
            outlook.to,
            vec![Person {
                name: "to@domain.com".to_string(),
                email: "to@domain.com".to_string()
            },]
        );

        assert_eq!(
            outlook.subject,
            String::from("creating an outlook message file")
        );
    }

    #[test]
    fn test_recipient_types() {
        let path = "data/test_email.msg";
        let outlook = Outlook::from_path(path).unwrap();

        // Verify recipient splitting by RecipientType
        assert_eq!(outlook.to.len(), 1);
        assert_eq!(outlook.cc.len(), 2);
        assert_eq!(outlook.bcc.len(), 3);

        // ascii.msg has only To recipients
        let outlook = Outlook::from_path("data/ascii.msg").unwrap();
        assert_eq!(outlook.to.len(), 1);
        assert!(outlook.cc.is_empty());
        assert!(outlook.bcc.is_empty());
    }

    #[test]
    fn test_date_fields() {
        // unicode.msg has all four date fields
        let outlook = Outlook::from_path("data/unicode.msg").unwrap();
        assert_eq!(outlook.client_submit_time, "2013-11-18T08:26:24Z");
        assert_eq!(outlook.message_delivery_time, "2013-11-18T08:26:29Z");
        assert!(outlook.creation_time.starts_with("2013-11-18T08:32:28"));
        assert!(
            outlook
                .last_modification_time
                .starts_with("2013-11-18T08:32:28")
        );

        // test_email.msg has delivery time but no submit time
        let outlook = Outlook::from_path("data/test_email.msg").unwrap();
        assert!(outlook.client_submit_time.is_empty());
        assert!(
            outlook
                .message_delivery_time
                .starts_with("2021-01-05T03:00:32")
        );
        assert!(outlook.creation_time.starts_with("2021-01-05T03:13:18"));

        // ascii.msg has no submit/delivery times
        let outlook = Outlook::from_path("data/ascii.msg").unwrap();
        assert!(outlook.client_submit_time.is_empty());
        assert!(outlook.message_delivery_time.is_empty());
        assert!(outlook.creation_time.starts_with("2017-06-01T15:24:31"));
    }

    #[test]
    fn test_message_class_importance_sensitivity() {
        let outlook = Outlook::from_path("data/test_email.msg").unwrap();
        assert_eq!(outlook.message_class, "IPM.Note");
        assert_eq!(outlook.importance, 1); // Normal
        assert_eq!(outlook.sensitivity, 0); // Normal

        let outlook = Outlook::from_path("data/unicode.msg").unwrap();
        assert_eq!(outlook.message_class, "IPM.Note");
        assert_eq!(outlook.importance, 1);
        // unicode.msg has no sensitivity property, defaults to 0
        assert_eq!(outlook.sensitivity, 0);

        let outlook = Outlook::from_path("data/ascii.msg").unwrap();
        assert_eq!(outlook.message_class, "IPM.Note");
    }

    #[test]
    fn test_from_reader() {
        let file = std::fs::File::open("data/unicode.msg").unwrap();
        let reader_outlook = Outlook::from_reader(file).unwrap();
        let path_outlook = Outlook::from_path("data/unicode.msg").unwrap();

        assert_eq!(reader_outlook.subject, path_outlook.subject);
        assert_eq!(reader_outlook.sender, path_outlook.sender);
        assert_eq!(reader_outlook.to, path_outlook.to);
        assert_eq!(reader_outlook.cc, path_outlook.cc);
    }

    #[test]
    fn test_to_json() {
        let path = "data/test_email.msg";
        let outlook = Outlook::from_path(path).unwrap();
        let json = outlook.to_json().unwrap();
        assert!(!json.is_empty());
    }

    #[test]
    fn test_html_field_present() {
        // test_email.msg has no Html property, so html should be empty
        let outlook = Outlook::from_path("data/test_email.msg").unwrap();
        assert!(outlook.html.is_empty());

        // unicode.msg also has no Html property
        let outlook = Outlook::from_path("data/unicode.msg").unwrap();
        assert!(outlook.html.is_empty());
    }

    #[test]
    fn test_html_in_json_output() {
        let outlook = Outlook::from_path("data/test_email.msg").unwrap();
        let json = outlook.to_json().unwrap();
        assert!(json.contains("\"html\""));
    }
}
