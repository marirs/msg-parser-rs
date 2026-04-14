use std::{fs::File, path::Path, sync::LazyLock};

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

// TransportHeaders contains transport specific message
// envelope information for the email.
#[non_exhaustive]
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct TransportHeaders {
    pub raw: String,
    pub content_type: String,
    pub date: String,
    pub message_id: String,
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

// Person represents either Sender or Receiver.
#[non_exhaustive]
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct Person {
    pub name: Name,
    pub email: Email,
}

impl Person {
    #[cfg(test)]
    fn new(name: Name, email: Email) -> Self {
        Self { name, email }
    }
    fn create_from_props(props: &Properties, name_key: &str, email_keys: &[&str]) -> Self {
        let name: String = props.get(name_key).map_or(String::new(), |x| x.into());
        // Get the fist email that can be found in props given email_keys.
        let email = email_keys
            .iter()
            .map(|&key| props.get(key).map_or(String::new(), |x| x.into()))
            .find(|x| !x.is_empty())
            .unwrap_or_default();
        Self { name, email }
    }
}

// Attachment represents attachment object in the mail.
#[non_exhaustive]
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct Attachment {
    pub display_name: String, // "DisplayName"
    pub payload: String,      // "AttachDataObject" (hex-encoded)
    #[serde(with = "hex")]
    pub payload_bytes: Vec<u8>, // "AttachDataObject" (raw bytes)
    pub extension: String,    // "AttachExtension"
    pub mime_tag: String,     // "AttachMimeTag"
    pub file_name: String,    // "AttachFilename" (8.3 short name)
    pub long_file_name: String, // "AttachLongFilename" (full name)
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
        }
    }
}

// Outlook is the Mail container.
// Each field corresponds to a field listed in
// MS-OXPROPS.
// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/f6ab1613-aefe-447d-a49c-18217230b148
// Note: Prefixes are omitted for brevity.
#[non_exhaustive]
#[derive(Serialize, Deserialize, Debug)]
pub struct Outlook {
    pub headers: TransportHeaders,      // "TransportMessageHeader"
    pub sender: Person,                 // "SenderName" , "SenderSmtpAddress"/"SenderEmailAddress"
    pub to: Vec<Person>,                // RecipientType == 1
    pub cc: Vec<Person>,                // RecipientType == 2
    pub bcc: Vec<Person>,               // RecipientType == 3
    pub subject: String,                // "Subject"
    pub body: String,                   // "Body"
    pub html: String,                   // "Html" (0x1013)
    pub rtf_compressed: String,         // "RtfCompressed"
    pub message_class: String,          // "MessageClass" (0x001A) e.g. "IPM.Note"
    pub importance: u32,                // "Importance" (0x0017) 0=Low, 1=Normal, 2=High
    pub sensitivity: u32, // "Sensitivity" (0x0036) 0=Normal, 1=Personal, 2=Private, 3=Confidential
    pub client_submit_time: String, // "ClientSubmitTime" (0x0039) ISO 8601 UTC
    pub message_delivery_time: String, // "MessageDeliveryTime" (0x0E06) ISO 8601 UTC
    pub creation_time: String, // "CreationTime" (0x3007) ISO 8601 UTC
    pub last_modification_time: String, // "LastModificationTime" (0x3008) ISO 8601 UTC
    pub attachments: Vec<Attachment>, // See Attachment struct
}

impl Outlook {
    fn populate(storages: &Storages) -> Self {
        let headers_text = storages.get_val_from_root_or_default("TransportMessageHeaders");
        let headers = TransportHeaders::create_from_headers_text(&headers_text);

        let mut to = Vec::new();
        let mut cc = Vec::new();
        let mut bcc = Vec::new();

        for (i, recip_map) in storages.recipients.iter().enumerate() {
            let person = Person::create_from_props(
                recip_map,
                "DisplayName",
                &["SmtpAddress", "EmailAddress"],
            );
            // RecipientType: 1=To, 2=CC, 3=BCC (MS-OXMSG 2.2.1)
            match storages.get_recipient_int_prop(i, "RecipientType") {
                Some(2) => cc.push(person),
                Some(3) => bcc.push(person),
                _ => to.push(person), // Default to To (including type==1 and missing)
            }
        }

        Self {
            headers,
            sender: Person::create_from_props(
                &storages.root,
                "SenderName",
                &["SenderSmtpAddress", "SenderEmailAddress"],
            ),
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

    pub fn from_path<P: AsRef<Path>>(path: P) -> Result<Self, Error> {
        let file = File::open(path)?;
        let parser = ole::Reader::new(file)?;
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);

        let outlook = Self::populate(&storages);
        Ok(outlook)
    }

    pub fn from_reader<R: std::io::Read>(mut reader: R) -> Result<Self, Error> {
        let mut buf = Vec::new();
        reader.read_to_end(&mut buf)?;
        Self::from_slice(&buf)
    }

    pub fn from_slice(slice: &[u8]) -> Result<Self, Error> {
        let parser = ole::Reader::new(slice)?;
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);

        let outlook = Self::populate(&storages);
        Ok(outlook)
    }

    pub fn to_json(&self) -> Result<String, Error> {
        Ok(serde_json::to_string(self)?)
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
            "Error parsing file with ole: failed to fill whole buffer".to_string()
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
