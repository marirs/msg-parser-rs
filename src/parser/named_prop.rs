//! Named property resolution for MAPI property IDs >= 0x8000.
//!
//! The `__nameid_version1.0` storage in an .msg file contains three streams
//! that map property IDs in the 0x8000+ range to named properties (either
//! GUID+dispID pairs or GUID+string name pairs).
//!
//! See: [MS-OXMSG] Section 2.2.3 and [MS-OXPROPS].

use std::collections::HashMap;

/// Well-known property set GUIDs (little-endian byte representation).
const PS_PUBLIC_STRINGS: [u8; 16] = [
    0x29, 0x03, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];
const PSETID_COMMON: [u8; 16] = [
    0x08, 0x20, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];
const PSETID_ADDRESS: [u8; 16] = [
    0x04, 0x20, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];
const PSETID_APPOINTMENT: [u8; 16] = [
    0x02, 0x20, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];
const PSETID_TASK: [u8; 16] = [
    0x03, 0x20, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];
const PSETID_LOG: [u8; 16] = [
    0x0A, 0x20, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];
const PSETID_NOTE: [u8; 16] = [
    0x0E, 0x20, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x46,
];

/// Well-known named properties: (GUID, dispID) -> canonical name.
fn well_known_dispid(guid: &[u8; 16], disp_id: u32) -> Option<&'static str> {
    match (guid, disp_id) {
        // PSETID_Common
        (&PSETID_COMMON, 0x8501) => Some("ReminderDelta"),
        (&PSETID_COMMON, 0x8502) => Some("ReminderTime"),
        (&PSETID_COMMON, 0x8503) => Some("ReminderSet"),
        (&PSETID_COMMON, 0x8506) => Some("Private"),
        (&PSETID_COMMON, 0x8510) => Some("SideEffects"),
        (&PSETID_COMMON, 0x8514) => Some("SmartNoAttach"),
        (&PSETID_COMMON, 0x8516) => Some("CommonStart"),
        (&PSETID_COMMON, 0x8517) => Some("CommonEnd"),
        (&PSETID_COMMON, 0x8520) => Some("FlagStatus"),
        (&PSETID_COMMON, 0x8530) => Some("FlagRequest"),
        (&PSETID_COMMON, 0x8535) => Some("TodoTitle"),
        (&PSETID_COMMON, 0x8560) => Some("ReminderSignalTime"),
        (&PSETID_COMMON, 0x8580) => Some("InternetAccountName"),
        (&PSETID_COMMON, 0x8581) => Some("InternetAccountStamp"),
        (&PSETID_COMMON, 0x8584) => Some("UseTnef"),
        (&PSETID_COMMON, 0x859C) => Some("InternetMessageId"),

        // PSETID_Address
        (&PSETID_ADDRESS, 0x8005) => Some("FileUnder"),
        (&PSETID_ADDRESS, 0x8006) => Some("FileUnderId"),
        (&PSETID_ADDRESS, 0x8080) => Some("Email1DisplayName"),
        (&PSETID_ADDRESS, 0x8082) => Some("Email1AddressType"),
        (&PSETID_ADDRESS, 0x8083) => Some("Email1Address"),
        (&PSETID_ADDRESS, 0x8084) => Some("Email1OriginalDisplayName"),
        (&PSETID_ADDRESS, 0x8085) => Some("Email1OriginalEntryId"),
        (&PSETID_ADDRESS, 0x8090) => Some("Email2DisplayName"),
        (&PSETID_ADDRESS, 0x8093) => Some("Email2Address"),
        (&PSETID_ADDRESS, 0x80A0) => Some("Email3DisplayName"),
        (&PSETID_ADDRESS, 0x80A3) => Some("Email3Address"),
        (&PSETID_ADDRESS, 0x80D8) => Some("InstantMessagingAddress"),

        // PSETID_Appointment
        (&PSETID_APPOINTMENT, 0x8205) => Some("BusyStatus"),
        (&PSETID_APPOINTMENT, 0x8208) => Some("Location"),
        (&PSETID_APPOINTMENT, 0x820D) => Some("AppointmentStartWhole"),
        (&PSETID_APPOINTMENT, 0x820E) => Some("AppointmentEndWhole"),
        (&PSETID_APPOINTMENT, 0x8213) => Some("AppointmentDuration"),
        (&PSETID_APPOINTMENT, 0x8215) => Some("AppointmentRecur"),
        (&PSETID_APPOINTMENT, 0x8216) => Some("AppointmentStateFlags"),
        (&PSETID_APPOINTMENT, 0x8217) => Some("ResponseStatus"),
        (&PSETID_APPOINTMENT, 0x8218) => Some("AppointmentReplyTime"),
        (&PSETID_APPOINTMENT, 0x8223) => Some("Recurring"),
        (&PSETID_APPOINTMENT, 0x8228) => Some("ExceptionReplaceTime"),
        (&PSETID_APPOINTMENT, 0x8231) => Some("AppointmentSubType"),
        (&PSETID_APPOINTMENT, 0x8232) => Some("AppointmentColor"),
        (&PSETID_APPOINTMENT, 0x8234) => Some("TimeZoneDescription"),
        (&PSETID_APPOINTMENT, 0x8235) => Some("TimeZoneStruct"),
        (&PSETID_APPOINTMENT, 0x8256) => Some("AllAttendeesString"),

        // PSETID_Task
        (&PSETID_TASK, 0x8101) => Some("TaskStatus"),
        (&PSETID_TASK, 0x8102) => Some("PercentComplete"),
        (&PSETID_TASK, 0x8103) => Some("TeamTask"),
        (&PSETID_TASK, 0x8104) => Some("TaskStartDate"),
        (&PSETID_TASK, 0x8105) => Some("TaskDueDate"),
        (&PSETID_TASK, 0x810F) => Some("TaskDateCompleted"),
        (&PSETID_TASK, 0x8110) => Some("TaskActualEffort"),
        (&PSETID_TASK, 0x8111) => Some("TaskEstimatedEffort"),
        (&PSETID_TASK, 0x811C) => Some("TaskComplete"),
        (&PSETID_TASK, 0x811F) => Some("TaskOwner"),
        (&PSETID_TASK, 0x8121) => Some("TaskAssigner"),
        (&PSETID_TASK, 0x8126) => Some("TaskFRecurring"),

        // PSETID_Log
        (&PSETID_LOG, 0x8700) => Some("LogType"),
        (&PSETID_LOG, 0x8706) => Some("LogStart"),
        (&PSETID_LOG, 0x8707) => Some("LogDuration"),
        (&PSETID_LOG, 0x8708) => Some("LogEnd"),

        // PSETID_Note (sticky notes)
        (&PSETID_NOTE, 0x8B00) => Some("NoteColor"),
        (&PSETID_NOTE, 0x8B02) => Some("NoteWidth"),
        (&PSETID_NOTE, 0x8B03) => Some("NoteHeight"),

        _ => None,
    }
}

/// Resolved named property: either a well-known canonical name or a string name
/// from the file itself.
#[derive(Debug, Clone)]
pub(crate) enum NamedPropName {
    Canonical(&'static str),
    Custom(String),
}

impl NamedPropName {
    pub(crate) fn as_str(&self) -> &str {
        match self {
            NamedPropName::Canonical(s) => s,
            NamedPropName::Custom(s) => s,
        }
    }
}

/// Maps property IDs (0x8000+) to their resolved names.
#[derive(Debug, Default)]
pub(crate) struct NamedPropertyMap {
    map: HashMap<u16, NamedPropName>,
}

impl NamedPropertyMap {
    /// Parse the three __nameid_version1.0 streams and build the mapping.
    pub(crate) fn parse(
        guid_stream: &[u8],
        entry_stream: &[u8],
        string_stream: &[u8],
    ) -> Self {
        let mut map = HashMap::new();

        if entry_stream.len() < 8 {
            return Self { map };
        }

        let num_entries = entry_stream.len() / 8;

        for i in 0..num_entries {
            let offset = i * 8;
            if offset + 8 > entry_stream.len() {
                break;
            }

            let name_id = u32::from_le_bytes([
                entry_stream[offset],
                entry_stream[offset + 1],
                entry_stream[offset + 2],
                entry_stream[offset + 3],
            ]);

            let packed = u32::from_le_bytes([
                entry_stream[offset + 4],
                entry_stream[offset + 5],
                entry_stream[offset + 6],
                entry_stream[offset + 7],
            ]);

            let kind = packed & 1;
            let guid_index = ((packed >> 1) & 0x7FFF) as usize;
            let prop_index = (packed >> 16) as u16;
            let prop_id = 0x8000 + prop_index;

            // Resolve the GUID
            let guid: Option<[u8; 16]> = match guid_index {
                0 => None, // PS_MAPI — not a named property
                1 => Some(PS_PUBLIC_STRINGS),
                2 => None, // PS_NONE
                n => {
                    let real_idx = n - 3;
                    let g_offset = real_idx * 16;
                    if g_offset + 16 <= guid_stream.len() {
                        let mut g = [0u8; 16];
                        g.copy_from_slice(&guid_stream[g_offset..g_offset + 16]);
                        Some(g)
                    } else {
                        None
                    }
                }
            };

            let Some(guid) = guid else {
                continue;
            };

            if kind == 0 {
                // Numeric (dispID) named property
                if let Some(name) = well_known_dispid(&guid, name_id) {
                    map.insert(prop_id, NamedPropName::Canonical(name));
                }
            } else {
                // String named property — read from string stream
                let str_offset = name_id as usize;
                if str_offset + 4 <= string_stream.len() {
                    let str_len = u32::from_le_bytes([
                        string_stream[str_offset],
                        string_stream[str_offset + 1],
                        string_stream[str_offset + 2],
                        string_stream[str_offset + 3],
                    ]) as usize;

                    let str_start = str_offset + 4;
                    if str_start + str_len <= string_stream.len() {
                        // UTF-16LE decode
                        let utf16_data = &string_stream[str_start..str_start + str_len];
                        let mut units = Vec::with_capacity(str_len / 2);
                        for chunk in utf16_data.chunks_exact(2) {
                            units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
                        }
                        let name = String::from_utf16_lossy(&units);
                        // Trim null terminators
                        let name = name.trim_end_matches('\0').to_string();
                        if !name.is_empty() {
                            map.insert(prop_id, NamedPropName::Custom(name));
                        }
                    }
                }
            }
        }

        Self { map }
    }

    /// Look up a property ID. Returns the name if it's a known named property.
    pub(crate) fn get(&self, prop_id: u16) -> Option<&str> {
        self.map.get(&prop_id).map(|n| n.as_str())
    }

    #[cfg(test)]
    pub(crate) fn is_empty(&self) -> bool {
        self.map.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ole::Reader;
    use std::io::Read;

    fn read_nameid_streams(path: &str) -> (Vec<u8>, Vec<u8>, Vec<u8>) {
        let parser = Reader::from_path(path).unwrap();
        let mut guid_stream = Vec::new();
        let mut entry_stream = Vec::new();
        let mut string_stream = Vec::new();

        // Find the __nameid_version1.0 storage ID
        let mut nameid_id = None;
        for entry in parser.iterate() {
            if entry.name() == "__nameid_version1.0" {
                nameid_id = Some(entry.id());
                break;
            }
        }

        if let Some(nid) = nameid_id {
            for entry in parser.iterate() {
                if entry.parent_node() == Some(nid) {
                    if let Ok(mut slice) = parser.get_entry_slice(entry) {
                        let mut buf = vec![0u8; slice.len()];
                        let _ = slice.read(&mut buf);
                        match entry.name() {
                            "__substg1.0_00020102" => guid_stream = buf,
                            "__substg1.0_00030102" => entry_stream = buf,
                            "__substg1.0_00040102" => string_stream = buf,
                            _ => {}
                        }
                    }
                }
            }
        }

        (guid_stream, entry_stream, string_stream)
    }

    #[test]
    fn test_parse_named_props_test_email() {
        let (guid, entry, string) = read_nameid_streams("data/test_email.msg");
        let map = NamedPropertyMap::parse(&guid, &entry, &string);
        assert!(!map.is_empty());
        // Print all resolved names for debugging
        let mut entries: Vec<_> = map.map.iter().collect();
        entries.sort_by_key(|&(id, _)| *id);
        for (id, name) in &entries {
            println!("  0x{:04X} -> {}", id, name.as_str());
        }
    }

    #[test]
    fn test_parse_named_props_unicode() {
        let (guid, entry, string) = read_nameid_streams("data/unicode.msg");
        let map = NamedPropertyMap::parse(&guid, &entry, &string);
        // unicode.msg has some named properties
        let mut entries: Vec<_> = map.map.iter().collect();
        entries.sort_by_key(|&(id, _)| *id);
        for (id, name) in &entries {
            println!("  0x{:04X} -> {}", id, name.as_str());
        }
    }

    #[test]
    fn test_empty_streams() {
        let map = NamedPropertyMap::parse(&[], &[], &[]);
        assert!(map.is_empty());
    }
}
