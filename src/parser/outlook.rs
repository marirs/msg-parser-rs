use std::{
    fs::File,
    path::Path
};

use regex::Regex;

use serde::{Deserialize, Serialize};
use serde_json;

use crate::ole;

use super::{
    error::Error,
    storage::{
        Properties,
        Storages
    }
};

type Name = String;
type Email = String;

// TransportHeaders contains transport specific message
// envelope information for the email.
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct TransportHeaders {
    pub content_type: String,
    pub date: String,
    pub message_id: String,
    pub reply_to: String,
}

impl TransportHeaders {
    fn extract_field(text: &str, re: Regex) -> String {
        if text.len() == 0 {
            return String::from("");
        }
        let caps = re.captures(text);
        if caps.is_none() {
            return String::from("");
        }
        caps.and_then(|cap| cap.get(1).map(|x| String::from(x.as_str())))
            .unwrap_or(String::from(""))
    }

    pub fn create_from_headers_text(text: &str) -> Self {
        // Case-insensitive match
        Self {
            content_type: Self::extract_field(
                text,
                Regex::new(r"(?i)Content-Type: (.*(\n\s.*)*)\r\n").unwrap(),
            ),
            date: Self::extract_field(&text, Regex::new(r"(?i)Date: (.*(\n\s.*)*)\r\n").unwrap()),
            message_id: Self::extract_field(
                text,
                Regex::new(r"(?i)Message-ID: (.*(\n\s.*)*)\r\n").unwrap(),
            ),
            reply_to: Self::extract_field(
                text,
                Regex::new(r"(?i)Reply-To: (.*(\n\s.*)*)\r\n").unwrap(),
            ),
        }
    }
}

// Person represents either Sender or Receiver.
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct Person {
    pub name: Name,
    pub email: Email,
}

impl Person {
    fn new(name: Name, email: Email) -> Self {
        Self { name, email }
    }
    fn create_from_props(props: &Properties, name_key: &str, email_keys: Vec<&str>) -> Self {
        let name: String = props.get(name_key).map_or(String::new(), |x| x.into());
        // Get the fist email that can be found in props given email_keys.
        let email = email_keys
            .iter()
            .map(|&key| props.get(key).map_or(String::new(), |x| x.into()))
            .find(|x| x.len() > 0)
            .unwrap_or(String::from(""));
        Self { name, email }
    }
}

// Attachment represents attachment object in the mail.
#[derive(Debug, PartialEq, Serialize, Deserialize)]
pub struct Attachment {
    pub display_name: String, // "DisplayName"
    pub payload: String,      // "AttachDataObject"
    pub extension: String,    // "AttachExtension"
    pub mime_tag: String,     // "AttachMimeTag"
    pub file_name: String,    // "AttachFilename"
}

impl Attachment {
    fn create(storages: &Storages, idx: usize) -> Self {
        Self {
            display_name: storages.get_val_from_attachment_or_default(idx, "DisplayName"),
            payload: storages.get_val_from_attachment_or_default(idx, "AttachDataObject"),
            extension: storages.get_val_from_attachment_or_default(idx, "AttachExtension"),
            mime_tag: storages.get_val_from_attachment_or_default(idx, "AttachMimeTag"),
            file_name: storages.get_val_from_attachment_or_default(idx, "AttachFilename"),
        }
    }
}

// Outlook is the Mail container.
// Each field corresponds to a field listed in
// MS-OXPROPS.
// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/f6ab1613-aefe-447d-a49c-18217230b148
// Note: Prefixes are omitted for brevity.
#[derive(Serialize, Deserialize, Debug)]
pub struct Outlook {
    pub headers: TransportHeaders,    // "TransportMessageHeader"
    pub sender: Person,               // "SenderName" , "SenderSmtpAddress"/"SenderEmailAddress"
    pub to: Vec<Person>,              // "DisplayName", "SmtpAddress"/"EmailAddress"
    pub cc: Vec<Person>,              // "DisplayCc"
    pub bcc: Name,                    // "DisplayBcc"
    pub subject: String,              // "Subject"
    pub body: String,                 // "Body"
    pub rtf_compressed: String,       // "RtfCompressed"
    pub attachments: Vec<Attachment>, // See Attachment struct
}

impl Outlook {
    fn extract_cc_from_headers(header_text: &str) -> Vec<Person> {
        // Format in header is:
        // CC: NAME <EMAIL>, NAME <EMAIL> \r\n
        let re = Regex::new(r"(?i)CC: .*(\r\n\t)?.*\r\n").unwrap();
        let caps = re.captures(header_text);
        if caps.is_none() {
            return vec![];
        }
        let cap = caps.unwrap().get(0).unwrap().as_str();
        // Remove first 3 chars
        // Split at ",", then trim and clean each string
        // We should be left with ["NAME <EMAIL", "NAME <EMAIL"]
        let cc_list = &cap[3..]
            .split(",")
            .map(|x| x.trim().replace('>', ""))
            .collect::<Vec<String>>();

        let mut cc_persons: Vec<Person> = vec![];
        for cc in cc_list.iter() {
            let name_email_pair: Vec<&str> = cc.split("<").map(|x| x.trim()).collect();
            let person = if name_email_pair.len() < 2 {
                // In the unlikely event that there's no email provided.
                Person::new(name_email_pair[0].to_string(), "".to_string())
            } else {
                Person::new(
                    name_email_pair[0].replace('"', ""),
                    name_email_pair[1].to_string(),
                )
            };
            cc_persons.push(person);
        }
        cc_persons
    }

    fn populate(storages: &Storages) -> Self {
        let headers_text = storages.get_val_from_root_or_default("TransportMessageHeaders");
        let headers = TransportHeaders::create_from_headers_text(&headers_text);

        // Outlook::extract_cc_from_headers(&headers_text);
        Self {
            headers,
            sender: Person::create_from_props(
                &storages.root,
                "SenderName",
                vec!["SenderSmtpAddress", "SenderEmailAddress"],
            ),
            to: storages
                .recipients
                .iter()
                .map(|recip_map| {
                    Person::create_from_props(
                        recip_map,
                        "DisplayName",
                        vec!["SmtpAddress", "EmailAddress"],
                    )
                })
                .collect(),
            cc: Outlook::extract_cc_from_headers(&headers_text),
            bcc: storages.get_val_from_root_or_default("DisplayBcc"),
            subject: storages.get_val_from_root_or_default("Subject"),
            body: storages.get_val_from_root_or_default("Body"),
            rtf_compressed: storages.get_val_from_root_or_default("RtfCompressed"),
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

        assert_eq!(
            header,
            TransportHeaders {
                content_type: String::new(),
                date: String::new(),
                message_id: String::new(),
                reply_to: String::new()
            }
        );
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
        assert_eq!(
            outlook.to,
            vec![
                Person {
                    name: "marirs@outlook.com".to_string(),
                    email: "marirs@outlook.com".to_string()
                },
                Person {
                    name: "Sriram Govindan".to_string(),
                    email: "marirs@aol.in".to_string()
                },
                Person {
                    name: "marirs@outlook.in".to_string(),
                    email: "marirs@outlook.in".to_string()
                },
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

        assert_eq!(
            outlook.subject,
            String::from("Test Email")
        );

        assert_eq!(
            outlook.headers,
            TransportHeaders {
                content_type: String::new(),
                date: String::new(),
                message_id: String::new(),
                reply_to: String::new(),
            }
        );

        assert_eq!(
            outlook
                .body
                .starts_with("Test Email\r\n"),
            true
        );
        assert_eq!(
            outlook.rtf_compressed.starts_with("51210000c8a200004c5a4"),
            true
        );
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
        assert_eq!(
            outlook.to,
            vec![
                Person {
                    name: "marirs@outlook.com".to_string(),
                    email: "marirs@outlook.com".to_string()
                },
                Person {
                    name: "Sriram Govindan".to_string(),
                    email: "marirs@aol.in".to_string()
                },
                Person {
                    name: "marirs@outlook.in".to_string(),
                    email: "marirs@outlook.in".to_string()
                },
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
        assert_eq!(
            outlook.subject,
            String::from("Test Email")
        );

        assert_eq!(
            outlook
                .body
                .starts_with("Test Email"),
            true
        );

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
        assert_eq!(
            mimes,
            vec![
                "".to_string(),
                "".to_string(),
                "".to_string()
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
        assert_eq!(
            outlook.to,
            vec![
                Person {
                    name: "brianzhou@me.com".to_string(),
                    email: "brianzhou@me.com".to_string()
                },
                Person {
                    name: "Brian Zhou".to_string(),
                    email: "brizhou@gmail.com".to_string(),
                }
            ]
        );

        assert_eq!(
            outlook.cc,
            vec![Person::new(
                "Brian Zhou".to_string(),
                "brizhou@gmail.com".to_string()
            ),]
        );
        assert_eq!(outlook.subject, String::from("Test for TIF files"));
        assert_eq!(
            outlook.headers,
            TransportHeaders {
                content_type: "multipart/mixed; boundary=001a113392ecbd7a5404eb6f4d6a".to_string(),
                date: "Mon, 18 Nov 2013 10:26:24 +0200".to_string(),
                message_id: "<CADtJ4eNjQSkGcBtVteCiTF+YFG89+AcHxK3QZ=-Mt48xygkvdQ@mail.gmail.com>"
                    .to_string(),
                reply_to: String::from("")
            }
        );
        assert_eq!(outlook.rtf_compressed.starts_with("bc020000b908"), true);
    }

    #[test]
    fn test_multiple_cc() {
        let path = "data/test_email.msg";
        let outlook = Outlook::from_path(path).unwrap();

        assert_eq!(
            outlook.cc,
            vec![]
        );
    }

    #[test]
    fn test_to_json() {
        let path = "data/test_email.msg";
        let outlook = Outlook::from_path(path).unwrap();
        let json = outlook.to_json().unwrap();
        assert_eq!(json.len() > 0, true);
    }
}
