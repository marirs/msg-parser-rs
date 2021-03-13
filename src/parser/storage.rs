use std::{
    collections::HashMap,
    u32::MAX,
};

use hex::decode;

use crate::ole::{Entry, EntryType, Reader};

use super::{
    constants::PropIdNameMap,
    decode::DataType,
    stream::Stream
};

// StorageType refers to major components in Message object.
// Refer to MS-OXPROPS 1.3.3
#[derive(Debug, Clone, PartialEq)]
pub enum StorageType {
    // u32 refers to its index
    Recipient(u32),
    // u32 refers to its index
    Attachment(u32),
    RootEntry,
}

impl StorageType {
    fn convert_id_to_u32(id: &str) -> Option<u32> {
        // id is 8 digits hexadecimal sequence.
        if id.len() != 8 {
            return None;
        }
        // [0, 0, 0, 0] where each item is of base 256 (16x16).
        let decoded = decode(id).ok()?;
        let mut base = 1u32;
        let mut sum = 0u32;
        for &num in decoded.iter().rev() {
            sum = sum + num as u32 * base;
            if base >= MAX / 256 {
                break;
            }
            base *= 256;
        }
        Some(sum)
    }

    pub fn create(name: &str) -> Option<Self> {
        if name.starts_with("__recip_version1.0_") {
            // Extract the digits after '#' in __recip_version1.0_#00000000
            // Remaining digits is the index of Recipient.
            let id = name.split("#").collect::<Vec<&str>>()[1];
            let id_as_num = StorageType::convert_id_to_u32(id)?;
            return Some(StorageType::Recipient(id_as_num));
        }
        if name.starts_with("__attach_version1.0_") {
            let id = name.split("#").collect::<Vec<&str>>()[1];
            let id_as_num = StorageType::convert_id_to_u32(id)?;
            return Some(StorageType::Attachment(id_as_num));
        }
        None
    }
}

// EntryStorageMap represents HashMap of ole::Entry id and its StorageType
#[derive(Debug)]
struct EntryStorageMap {
    map: HashMap<u32, StorageType>,
}

impl EntryStorageMap {
    pub fn new(parser: &Reader) -> Self {
        let mut storage_map: HashMap<u32, StorageType> = HashMap::new();
        for entry in parser.iterate() {
            match entry._type() {
                EntryType::RootStorage => {
                    storage_map.insert(entry.id(), StorageType::RootEntry);
                }
                EntryType::UserStorage => {
                    StorageType::create(entry.name())
                        .and_then(|storage| storage_map.insert(entry.id(), storage));
                }
                _ => {
                    continue;
                }
            }
        }
        Self { map: storage_map }
    }

    pub fn get_storage_type(&self, parent_id: Option<u32>) -> Option<&StorageType> {
        self.map.get(&parent_id?)
    }
}

// Properties is a Map is a collection of Message object elements.
pub type Properties = HashMap<String, DataType>;

// Recipients represent array of Recipient objects in Message.
pub type Recipients = Vec<Properties>;

// Attachments represent array of Attachment object in Message
pub type Attachments = Vec<Properties>;

// Storages is a collection of Storage
// object containing their decoded stream
// values for respective properties.
#[derive(Debug)]
pub struct Storages {
    storage_map: EntryStorageMap,
    prop_map: PropIdNameMap,
    pub attachments: Attachments,
    pub recipients: Recipients,
    // Mail properties
    pub root: Properties,
}

impl Storages {
    fn to_arr(map: HashMap<u32, Properties>) -> Vec<Properties> {
        let mut tuples: Vec<(u32, Properties)> = map
            .into_iter()
            .map(|(k, v)| (k, v))
            .collect::<Vec<(u32, Properties)>>();
        tuples.sort_by(|a, b| a.0.cmp(&b.0));
        tuples.into_iter().map(|x| x.1).collect::<Vec<Properties>>()
    }

    fn create_stream(&self, parser: &Reader, entry: &Entry) -> Option<Stream> {
        let parent = self.storage_map.get_storage_type(entry.parent_node())?;
        let mut slice = parser.get_entry_slice(entry).ok()?;
        Stream::create(entry.name(), &mut slice, &self.prop_map, parent)
    }

    pub fn process_streams(&mut self, parser: &Reader) {
        let mut recipients_map: HashMap<u32, Properties> = HashMap::new();
        let mut attachments_map: HashMap<u32, Properties> = HashMap::new();
        for entry in parser.iterate() {
            if let EntryType::UserStream = entry._type() {
                // Decode stream from slice.
                // Skip if failed.
                let stream_res = self.create_stream(&parser, &entry);
                if stream_res.is_none() {
                    continue;
                }
                let stream = stream_res.unwrap();

                // Populate maps accordingly
                match stream.parent {
                    StorageType::RootEntry => {
                        self.root.insert(stream.key, stream.value);
                    }
                    StorageType::Recipient(id) => {
                        let recipient_map = recipients_map.entry(id).or_insert(HashMap::new());
                        (*recipient_map).insert(stream.key, stream.value);
                    }
                    StorageType::Attachment(id) => {
                        let attachment_map = attachments_map.entry(id).or_insert(HashMap::new());
                        (*attachment_map).insert(stream.key, stream.value);
                    }
                }
            }
        }
        // Update storages
        self.recipients = Self::to_arr(recipients_map);
        self.attachments = Self::to_arr(attachments_map);
    }

    pub fn new(parser: &Reader) -> Self {
        let root: Properties = HashMap::new();
        let recipients: Recipients = vec![];
        let attachments: Attachments = vec![];
        let storage_map = EntryStorageMap::new(parser);
        let prop_map = PropIdNameMap::init();
        Self {
            storage_map,
            prop_map,
            root,
            recipients,
            attachments,
        }
    }

    pub fn get_val_from_root_or_default(&self, key: &str) -> String {
        self.root.get(key).map_or(String::new(), |x| x.into())
    }

    pub fn get_val_from_attachment_or_default(&self, idx: usize, key: &str) -> String {
        self.attachments
            .iter()
            .nth(idx)
            .map(|attach| attach.get(key).map_or(String::from(""), |x| x.into()))
            .unwrap_or(String::new())
    }
}

#[cfg(test)]
mod tests {
    use super::super::decode::DataType;
    use super::{EntryStorageMap, Properties, StorageType, Storages};
    use crate::ole::Reader;
    use std::collections::HashMap;

    #[test]
    fn test_storage_type_convert() {
        use std::u32::MAX;
        let mut id = StorageType::convert_id_to_u32("00000001");
        assert_eq!(id, Some(1u32));

        id = StorageType::convert_id_to_u32("0000000A");
        assert_eq!(id, Some(10u32));

        id = StorageType::convert_id_to_u32("00000101");
        assert_eq!(id, Some(257u32));

        id = StorageType::convert_id_to_u32("FFFFFFFF");
        assert_eq!(id, Some(MAX));

        // Edge Cases
        id = StorageType::convert_id_to_u32("HELLO");
        assert_eq!(id, None);

        id = StorageType::convert_id_to_u32("00000000000000");
        assert_eq!(id, None);
    }

    #[test]
    fn test_create_storage_type() {
        let recipient = StorageType::create("__recip_version1.0_#0000000A");
        assert_eq!(recipient, Some(StorageType::Recipient(10)));

        let attachment = StorageType::create("__attach_version1.0_#0000000A");
        assert_eq!(attachment, Some(StorageType::Attachment(10)));

        let unknown_storage = StorageType::create("");
        assert_eq!(unknown_storage, None);
    }

    #[test]
    fn test_storage_map() {
        let parser = Reader::from_path("data/test_email.msg").unwrap();
        let storage_map = EntryStorageMap::new(&parser);

        let mut expected_map = HashMap::new();
        expected_map.insert(0, StorageType::RootEntry);
        expected_map.insert(73, StorageType::Recipient(0));
        expected_map.insert(85, StorageType::Recipient(1));
        expected_map.insert(97, StorageType::Recipient(2));
        expected_map.insert(108, StorageType::Recipient(3));
        expected_map.insert(120, StorageType::Recipient(4));
        expected_map.insert(132, StorageType::Recipient(5));
        expected_map.insert(143, StorageType::Attachment(0));
        expected_map.insert(260, StorageType::Recipient(0));
        expected_map.insert(310, StorageType::Attachment(1));
        expected_map.insert(323, StorageType::Attachment(2));
        assert_eq!(storage_map.map, expected_map);
    }

    #[test]
    fn test_storage_to_arr() {
        let mut map_apple: Properties = HashMap::new();
        map_apple.insert("A".to_string(), DataType::PtypString("Apple".to_string()));
        let mut map_bagel: Properties = HashMap::new();
        map_bagel.insert("B".to_string(), DataType::PtypString("Bagel".to_string()));

        let mut basket: HashMap<u32, Properties> = HashMap::new();
        basket.insert(1, map_apple);
        basket.insert(0, map_bagel);

        let res = Storages::to_arr(basket);
        assert_eq!(
            res[0].get("B"),
            Some(&DataType::PtypString("Bagel".to_string()))
        );
        assert_eq!(
            res[1].get("A"),
            Some(&DataType::PtypString("Apple".to_string()))
        );
    }

    #[test]
    fn test_create_storage_test_email() {
        let parser = Reader::from_path("data/test_email.msg").unwrap();
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);

        let sender = storages.root.get("SenderEmailAddress");
        assert!(sender.is_none());

        // Check attachments
        assert_eq!(storages.attachments.len(), 3);

        // Check recipients
        assert_eq!(storages.recipients.len(), 6);

        // Check Display name
        let display_name = storages.recipients[0].get("DisplayName").unwrap();
        assert_eq!(
            display_name,
            &DataType::PtypString("marirs@outlook.com".to_string())
        );
    }

    #[test]
    fn test_create_storage_outlook_attachments() {
        let parser = Reader::from_path("data/test_email.msg").unwrap();
        let mut storages = Storages::new(&parser);
        storages.process_streams(&parser);


        // Check attachment
        assert_eq!(storages.attachments.len(), 3);

        let attachment_name = storages.attachments[0].get("DisplayName");
        assert_eq!(
            attachment_name,
            Some(&DataType::PtypString("1 Days Left\u{14} 35% off cloud space, upgrade now!".to_string()))
        );

        let attachment_name = storages.attachments[1].get("AttachFilename");
        assert_eq!(
            attachment_name,
            Some(&DataType::PtypString("milky-~1.jpg".to_string()))
        );

        let attachment_name = storages.attachments[2].get("AttachFilename");
        assert_eq!(
            attachment_name,
            Some(&DataType::PtypString("TestEm~1.msg".to_string()))
        );

        // Check recipients
        assert_eq!(storages.recipients.len(), 6);
        let display_name = storages.recipients[1].get("DisplayName").unwrap();
        assert_eq!(display_name, &DataType::PtypString("Sriram Govindan".to_string()));
    }
}
