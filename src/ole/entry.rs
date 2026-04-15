use crate::ole::util::FromSlice;

#[derive(Debug)]
pub(crate) enum NodeColour {
    Red,
    Black,
}

impl NodeColour {
    fn from(t: u8) -> Result<NodeColour, super::error::Error> {
        match t {
            0 => Ok(NodeColour::Red),
            1 => Ok(NodeColour::Black),
            _ => Err(super::error::Error::NodeTypeUnknown),
        }
    }
}

impl std::fmt::Display for NodeColour {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match *self {
            NodeColour::Red => write!(f, "RED"),
            NodeColour::Black => write!(f, "BLACK"),
        }
    }
}

#[derive(PartialEq, Debug, Clone, Copy)]
pub enum EntryType {
    /// Empty entry.
    Empty,

    /// Storage, i.e. a directory.
    UserStorage,

    /// Stream, i.e. a file.
    UserStream,

    /// LockBytes (unknown usage).
    LockBytes,

    /// Property (unknown usage).
    Property,

    /// Root storage.
    RootStorage,
}

impl EntryType {
    fn from(t: u8) -> Result<EntryType, super::error::Error> {
        match t {
            0 => Ok(EntryType::Empty),
            1 => Ok(EntryType::UserStorage),
            2 => Ok(EntryType::UserStream),
            3 => Ok(EntryType::LockBytes),
            4 => Ok(EntryType::Property),
            5 => Ok(EntryType::RootStorage),
            _ => Err(super::error::Error::NodeTypeUnknown),
        }
    }
}

impl std::fmt::Display for EntryType {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match *self {
            EntryType::Empty => write!(f, "Empty"),
            EntryType::UserStorage => write!(f, "User storage"),
            EntryType::UserStream => write!(f, "User stream"),
            EntryType::LockBytes => write!(f, "?? Lock bytes ??"),
            EntryType::Property => write!(f, "?? Property ??"),
            EntryType::RootStorage => write!(f, "Root storage"),
        }
    }
}

/// An entry in an OLE compound document.
///
/// An entry is either a *stream* (file) or a *storage* (directory).
///
/// # Example
///
/// ```rust,ignore
/// use crate::ole::Reader;
///
/// let reader = Reader::from_path("data/test_email.msg").unwrap();
/// let entry = reader.iterate().next().unwrap();
/// println!("Name: {}", entry.name());
/// println!("Type: {}", entry._type());
/// println!("Size: {}", entry.len());
/// ```
#[derive(Debug)]
pub struct Entry {
    /// ID of the entry.
    id: u32,

    /// Name of the stream or the storage.
    name: std::string::String,

    /// Type of the entry.
    entry_type: EntryType,

    /// Color of the entry (see <https://en.wikipedia.org/wiki/Red%E2%80%93black_tree>)
    color: NodeColour,

    /// ID of the left child entry.
    left_child_node: u32,

    /// ID of the right child entry.
    right_child_node: u32,

    /// ID of the root node
    root_node: u32,

    /// UID of the entry.
    identifier: std::vec::Vec<u8>, // 16 bytes

    /// Flags of the entry.
    flags: std::vec::Vec<u8>, // 4 bytes

    /// Creation time.
    creation_time: u64,

    /// Last modification time.
    last_modification_time: u64,

    /// Chain of secID which hold the stream or the storage
    sec_id_chain: std::vec::Vec<u32>,

    /// Size of the entry.
    size: usize,

    /// Array of the children's DirIDs
    children_nodes: std::vec::Vec<u32>,

    /// DirID of the parent
    parent_node: Option<u32>,
}

impl Entry {
    fn from_slice(sector: &[u8], dir_id: u32) -> Result<Entry, super::error::Error> {
        let entry = Entry {
            id: dir_id,
            name: Entry::build_name(&sector[0..64]),
            entry_type: EntryType::from(sector[66])?,
            color: NodeColour::from(sector[67])?,
            left_child_node: u32::from_slice(&sector[68..72]),
            right_child_node: u32::from_slice(&sector[72..76]),
            root_node: u32::from_slice(&sector[76..80]),
            identifier: sector[80..96].to_vec(),
            flags: sector[96..100].to_vec(),
            creation_time: u64::from_slice(&sector[100..108]),
            last_modification_time: u64::from_slice(&sector[108..116]),
            sec_id_chain: vec![u32::from_slice(&sector[116..120])],
            size: usize::from_slice(&sector[120..124]),
            children_nodes: std::vec::Vec::new(),
            parent_node: None,
        };

        Ok(entry)
    }

    fn build_name(array: &[u8]) -> std::string::String {
        // Entry names are UTF-16LE with a null terminator
        let mut units = Vec::with_capacity(32);
        let mut i = 0usize;
        while i + 1 < 64 && i + 1 < array.len() {
            let unit = u16::from_le_bytes([array[i], array[i + 1]]);
            if unit == 0 {
                break;
            }
            units.push(unit);
            i += 2;
        }
        String::from_utf16_lossy(&units)
    }

    /// Returns the ID of the entry.
    pub fn id(&self) -> u32 {
        self.id
    }

    /// Returns the creation time of the entry (could be 0)
    pub fn creation_time(&self) -> u64 {
        self.creation_time
    }

    /// Returns the last modification time of the entry (could be 0)
    pub fn last_modification_time(&self) -> u64 {
        self.last_modification_time
    }

    /// Returns the name of the entry.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the type of the entry.
    pub fn _type(&self) -> EntryType {
        self.entry_type
    }

    /// Returns the size of the entry
    pub fn len(&self) -> usize {
        self.size
    }

    /// Returns the DirID of the left child node
    pub fn left_child_node(&self) -> u32 {
        self.left_child_node
    }

    /// Returns the DirID of the right child node
    pub fn right_child_node(&self) -> u32 {
        self.right_child_node
    }

    /// Returns the DirID of the parent, if exists
    pub fn parent_node(&self) -> Option<u32> {
        self.parent_node
    }

    /// Returns the DirIDs of the children, if exists
    pub fn children_nodes(&self) -> &std::vec::Vec<u32> {
        &self.children_nodes
    }
}

impl std::fmt::Display for Entry {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(
            f,
            "Entry #{}. Type: {}, Color: {}, Name: {},
      Size: {}. SecID chain: {:?}",
            self.id, self.entry_type, self.color, &self.name, self.size, self.sec_id_chain
        )
    }
}

/// A readable slice of an entry's content.
///
/// OLE files use a FAT-like sector allocation scheme, so a stream can be
/// fragmented across non-contiguous sectors. `EntrySlice` reassembles these
/// chunks and implements [`Read`](std::io::Read).
///
/// # Example
///
/// ```rust,ignore
/// use crate::ole::Reader;
/// use std::io::Read;
///
/// let reader = Reader::from_path("data/test_email.msg").unwrap();
/// let entry = reader.iterate().next().unwrap();
/// let mut slice = reader.get_entry_slice(entry).unwrap();
/// let mut buf = [0u8; 42];
/// let n = slice.read(&mut buf).unwrap();
/// ```
pub struct EntrySlice<'s> {
    /// Chunk size, i.e. size of the sector.
    max_chunk_size: usize,

    /// List of slices.
    chunks: std::vec::Vec<&'s [u8]>,

    /// How many bytes which have been already read.
    read: usize,

    /// Total size of slice.
    total_size: usize,

    /// Real size of all chunks
    real_size: usize,
}

impl<'s> EntrySlice<'s> {
    fn new(max_chunk_size: usize, size: usize) -> EntrySlice<'s> {
        EntrySlice {
            max_chunk_size,
            chunks: std::vec::Vec::new(),
            read: 0usize,
            total_size: size,
            real_size: 0,
        }
    }

    fn add_chunk(&mut self, chunk: &'s [u8]) {
        self.real_size += chunk.len();
        self.chunks.push(chunk);
    }

    /// Returns the length of the slice, therefore the length of the entry.
    pub fn len(&self) -> usize {
        self.total_size
    }

    /// Returns the real length of all chunks
    pub fn real_len(&self) -> usize {
        self.real_size
    }
}

impl<'s> std::io::Read for EntrySlice<'s> {
    fn read(&mut self, buf: &mut [u8]) -> Result<usize, std::io::Error> {
        let to_read = std::cmp::min(buf.len(), self.total_size - self.read);
        if to_read == 0 || self.max_chunk_size == 0 {
            return Ok(0);
        }
        let mut offset = self.read;
        let mut read = 0;
        while read != to_read {
            let chunk_index = offset / self.max_chunk_size;
            if chunk_index >= self.chunks.len() {
                break;
            }
            let chunk = &self.chunks[chunk_index];
            let local_offset = offset % self.max_chunk_size;
            let end = std::cmp::min(local_offset + to_read - read, self.max_chunk_size);
            let n = end - local_offset;
            buf[read..read + n].copy_from_slice(&chunk[local_offset..end]);
            read += n;
            self.read += n;
            offset = self.read;
        }
        Ok(read)
    }
}

impl<'ole> super::ole::Reader<'ole> {
    /// Returns the slice for the entry.
    pub fn get_entry_slice(&self, entry: &Entry) -> Result<EntrySlice<'_>, super::error::Error> {
        let entry_slice: EntrySlice;
        let size = entry.size;
        if size == 0 {
            Err(super::error::Error::EmptyEntry)
        } else {
            if &size < self.minimum_standard_stream_size.as_ref().unwrap() {
                entry_slice = self.get_short_stream_slices(&entry.sec_id_chain, size)?;
            } else {
                entry_slice = self.get_stream_slices(&entry.sec_id_chain, size)?;
            }
            Ok(entry_slice)
        }
    }

    pub(crate) fn build_directory_entries(&mut self) -> Result<(), super::error::Error> {
        let n_entry_by_sector =
            self.sec_size.as_ref().unwrap() / super::constants::DIRECTORY_ENTRY_SIZE;
        let mut entries = std::vec::Vec::<Entry>::with_capacity(
            self.dsat.as_ref().unwrap().len() * n_entry_by_sector,
        );

        let mut k = 0usize;
        for i in 0..self.dsat.as_ref().unwrap().len() {
            let sector_index = self.dsat.as_ref().unwrap()[i];
            let sector = self.read_sector(sector_index as usize)?;
            for l in 0..n_entry_by_sector {
                let entry = Entry::from_slice(
                    &sector[l * super::constants::DIRECTORY_ENTRY_SIZE
                        ..(l + 1) * super::constants::DIRECTORY_ENTRY_SIZE],
                    k as u32,
                )?;
                entries.push(entry);
                k += 1;
            }
        }
        let stream_size = *self.minimum_standard_stream_size.as_ref().unwrap();
        for (i, entry) in entries.iter_mut().enumerate() {
            match entry.entry_type {
                EntryType::UserStream => {
                    let start_index = entry.sec_id_chain.pop().unwrap();
                    if entry.size < stream_size {
                        entry.sec_id_chain = self.build_chain_from_ssat(start_index);
                    } else {
                        entry.sec_id_chain = self.build_chain_from_sat(start_index);
                    }
                }
                EntryType::RootStorage => {
                    self.root_entry = Some(i as u32);
                    let start_index = entry.sec_id_chain.pop().unwrap();
                    entry.sec_id_chain = self.build_chain_from_sat(start_index);
                }
                _ => {}
            }
        }
        self.entries = Some(entries);
        self.build_entry_tree(0, None);
        Ok(())
    }

    fn get_short_stream_slices(
        &self,
        chain: &[u32],
        size: usize,
    ) -> Result<EntrySlice<'_>, super::error::Error> {
        let ssector_size = *self.short_sec_size.as_ref().unwrap();
        let mut entry_slice = EntrySlice::new(ssector_size, size);
        let entries = self
            .entries
            .as_ref()
            .ok_or(super::error::Error::BadSizeValue("No directory entries"))?;
        if entries.is_empty() {
            return Err(super::error::Error::BadSizeValue("No root entry"));
        }
        let short_stream_chain = &entries[0].sec_id_chain;
        let n_per_sector = *self.sec_size.as_ref().unwrap() / ssector_size;
        if n_per_sector == 0 {
            return Err(super::error::Error::BadSizeValue(
                "Invalid sector/short-sector size ratio",
            ));
        }
        let mut total_read = 0;
        for ssector_id in chain {
            let chain_idx = *ssector_id as usize / n_per_sector;
            if chain_idx >= short_stream_chain.len() {
                return Err(super::error::Error::BadSizeValue(
                    "Short sector ID out of range",
                ));
            }
            let sector_index = short_stream_chain[chain_idx];
            let sector = self.read_sector(sector_index as usize)?;
            let ssector_index = *ssector_id as usize % n_per_sector;
            let start = ssector_index * ssector_size;
            let end = start + std::cmp::min(ssector_size, size - total_read);
            if end > sector.len() {
                return Err(super::error::Error::BadSizeValue(
                    "Short sector slice out of range",
                ));
            }
            entry_slice.add_chunk(&sector[start..end]);
            total_read += end - start;
        }
        Ok(entry_slice)
    }

    fn get_stream_slices(
        &self,
        chain: &[u32],
        size: usize,
    ) -> Result<EntrySlice<'_>, super::error::Error> {
        let sector_size = *self.sec_size.as_ref().unwrap();
        let mut entry_slice = EntrySlice::new(sector_size, size);
        let mut total_read = 0;
        for sector_id in chain {
            let sector = self.read_sector(*sector_id as usize)?;
            let start = 0usize;
            let end = std::cmp::min(sector_size, size - total_read);
            entry_slice.add_chunk(&sector[start..end]);
            total_read += end - start;
        }
        Ok(entry_slice)
    }

    fn build_entry_tree(&mut self, id: u32, parent_id: Option<u32>) {
        let n = self.entries.as_ref().unwrap().len() as u32;
        let mut visited = std::collections::HashSet::new();
        // Stack items: (node_id, parent_id)
        let mut stack: Vec<(u32, Option<u32>)> = vec![(id, parent_id)];

        while let Some((current_id, current_parent)) = stack.pop() {
            if current_id == super::constants::FREE_SECID_U32 || current_id >= n {
                continue;
            }
            if !visited.insert(current_id) {
                // Already visited — skip to prevent cycles
                continue;
            }

            // Register the parent id for the current node
            self.entries.as_mut().unwrap()[current_id as usize].parent_node = current_parent;

            // Register as child
            if let Some(pid) = current_parent {
                self.entries.as_mut().unwrap()[pid as usize]
                    .children_nodes
                    .push(current_id);
            }

            let node_type = self.entries.as_ref().unwrap()[current_id as usize]._type();
            let left_child = self.entries.as_ref().unwrap()[current_id as usize].left_child_node();
            let right_child =
                self.entries.as_ref().unwrap()[current_id as usize].right_child_node();

            // Push right first so left is processed first (stack is LIFO)
            if right_child < n {
                stack.push((right_child, current_parent));
            }
            if left_child < n {
                stack.push((left_child, current_parent));
            }

            if node_type == EntryType::RootStorage || node_type == EntryType::UserStorage {
                let child = self.entries.as_ref().unwrap()[current_id as usize].root_node;
                stack.push((child, Some(current_id)));
            }
        }
    }
}
