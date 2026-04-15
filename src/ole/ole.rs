/// An OLE compound document reader.
///
/// The parsing method follows the specification described here:
/// <http://www.openoffice.org/sc/compdocfileformat.pdf>
///
/// # Example
///
/// ```rust,ignore
/// use crate::ole::Reader;
///
/// let reader = Reader::from_path("data/test_email.msg").unwrap();
/// for entry in reader.iterate() {
///     println!("{}", entry);
/// }
/// ```
pub struct Reader<'ole> {
    /// Buffer for reading from the source.
    pub(crate) buf_reader: Option<std::io::BufReader<Box<dyn std::io::Read + 'ole>>>,

    /// Unique identifier.
    pub(crate) uid: std::vec::Vec<u8>,

    /// Revision number.
    pub(crate) revision_number: Option<u16>,

    /// Version number.
    pub(crate) version_number: Option<u16>,

    /// Size of one sector.
    pub(crate) sec_size: Option<usize>,

    /// Size of one short sector.
    pub(crate) short_sec_size: Option<usize>,

    /// Sector Allocation Table.
    pub(crate) sat: Option<std::vec::Vec<u32>>,

    /// Directory Sector Allocation Table.
    pub(crate) dsat: Option<std::vec::Vec<u32>>,

    /// Minimum size of a standard stream size.
    pub(crate) minimum_standard_stream_size: Option<usize>,

    /// Short Sector Allocation Table.
    pub(crate) ssat: Option<std::vec::Vec<u32>>,

    /// Master Sector Allocation Table.
    pub(crate) msat: Option<std::vec::Vec<u32>>,

    /// Body of the file.
    pub(crate) body: Option<std::vec::Vec<u8>>,

    /// Directory entries.
    pub(crate) entries: Option<std::vec::Vec<super::entry::Entry>>,

    /// DirID of the root entry.
    pub(crate) root_entry: Option<u32>,
}

impl<'ole> Reader<'ole> {
    /// Constructs a new `Reader` from any [`Read`](std::io::Read) source.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use crate::ole::Reader;
    /// let file = std::fs::File::open("data/test_email.msg").unwrap();
    /// let reader = Reader::new(file).unwrap();
    /// ```
    pub fn new<T>(readable: T) -> std::result::Result<Reader<'ole>, super::error::Error>
    where
        T: std::io::Read + 'ole,
    {
        let mut t = Reader {
            buf_reader: Some(std::io::BufReader::new(Box::new(readable))),
            uid: vec![0u8; super::constants::UID_SIZE],
            revision_number: None,
            version_number: None,
            sec_size: None,
            short_sec_size: None,
            sat: None,
            dsat: None,
            minimum_standard_stream_size: None,
            ssat: None,
            msat: None,
            body: None,
            entries: None,
            root_entry: None,
        };
        t.parse_header()?;
        t.build_sat()?;
        t.build_directory_entries()?;
        Ok(t)
    }

    /// Constructs a new `Reader` from a file path.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use crate::ole::Reader;
    /// let reader = Reader::from_path("data/test_email.msg").unwrap();
    /// ```
    pub fn from_path(path: &str) -> Result<Reader<'_>, super::error::Error> {
        let data = std::fs::read(path).map_err(super::error::Error::IOError)?;
        Reader::from_bytes(data)
    }

    /// Constructs a new `Reader` from a byte vector or slice already in memory.
    ///
    /// This avoids the double-copy that occurs when passing a `&[u8]` to
    /// [`new()`](Reader::new), since `new()` wraps the source in a `BufReader`
    /// and then copies everything into an internal buffer via `read_to_end`.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use crate::ole::Reader;
    /// let data = std::fs::read("data/test_email.msg").unwrap();
    /// let reader = Reader::from_bytes(data).unwrap();
    /// ```
    pub fn from_bytes(data: impl Into<Vec<u8>>) -> Result<Reader<'ole>, super::error::Error> {
        let data = data.into();
        let mut t = Reader {
            buf_reader: None,
            uid: vec![0u8; super::constants::UID_SIZE],
            revision_number: None,
            version_number: None,
            sec_size: None,
            short_sec_size: None,
            sat: None,
            dsat: None,
            minimum_standard_stream_size: None,
            ssat: None,
            msat: None,
            body: None,
            entries: None,
            root_entry: None,
        };
        t.parse_header_from_bytes(&data)?;
        t.build_sat()?;
        t.build_directory_entries()?;
        Ok(t)
    }

    /// Returns an iterator over the directory entries of the OLE file.
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// use crate::ole::Reader;
    /// let reader = Reader::from_path("data/test_email.msg").unwrap();
    /// for entry in reader.iterate() {
    ///     println!("Entry: {} ({})", entry.name(), entry._type());
    /// }
    /// ```
    pub fn iterate(&self) -> super::iterator::OLEIterator<'_> {
        super::iterator::OLEIterator::new(self)
    }

    /// Read some bytes from the source.
    pub(crate) fn read(&mut self, buf: &mut [u8]) -> Result<usize, super::error::Error> {
        use std::io::Read;
        self.buf_reader
            .as_mut()
            .unwrap()
            .read_exact(buf)
            .map_err(super::error::Error::IOError)?;
        Ok(buf.len())
    }
}

#[cfg(test)]
mod tests {
    use super::super::error::Error;
    use super::Reader;

    #[test]
    fn instance_nok() {
        let path = "Thumbs.db";
        let o: Result<Reader, Error> = Reader::from_path(path);
        assert!(o.is_err());
        let e = o.err().unwrap();
        println!("NOK: {}", e);
    }

    #[test]
    fn instance_ok() {
        let path = "data/Thumbs.db";
        let o: Result<Reader, Error> = Reader::from_path(path);
        assert!(o.is_ok());
    }

    #[test]
    fn sector_sizes() {
        let ole: Reader = Reader::from_path("data/Thumbs.db").unwrap();
        assert_eq!(ole.sec_size, Some(512));
        assert_eq!(ole.short_sec_size, Some(64));
    }

    #[test]
    fn array_bad_identifier() {
        let mut vec = super::super::constants::IDENTIFIER.to_vec();
        vec[0] = 0xD1;
        fill(&mut vec);
        let ole = Reader::new(&vec[..]);
        assert!(ole.is_err());
        println!("BAD IDENTIFIER: {}", ole.err().unwrap());
    }

    fn fill(buf: &mut std::vec::Vec<u8>) {
        let missing = vec![0u8; super::super::constants::HEADER_SIZE - buf.len()];
        buf.extend(missing);
    }

    #[test]
    fn array_bad_endianness_identifier() {
        let mut vec = super::super::constants::IDENTIFIER.to_vec();
        vec.extend(vec![0u8; 20]);
        vec.push(0xFE);
        vec.push(0xFE);
        fill(&mut vec);
        let ole = Reader::new(&vec[..]);
        assert!(ole.is_err());
        println!("BAD ENDIANNESS: {}", ole.err().unwrap());
    }

    #[test]
    fn uid() {
        let ole = Reader::from_path("data/Thumbs.db");
        assert!(ole.is_ok());
        let ole = ole.unwrap();
        assert_eq!(&[0x0u8; 16], &ole.uid[..]);
    }

    #[test]
    fn bad_sec_size() {
        let mut vec = super::super::constants::IDENTIFIER.to_vec();
        vec.extend(vec![0x42u8; 20]);
        vec.extend(&super::super::constants::LITTLE_ENDIAN_IDENTIFIER);
        vec.extend(vec![0xFF, 0xFF, 0xFF, 0xFF]);
        vec.extend(vec![0u8; 10]);
        vec.extend(vec![0xFF, 0xFF, 0xFF, 0xFF]);
        fill(&mut vec);
        let ole = Reader::new(&vec[..]);
        assert!(ole.is_err());
    }

    #[test]
    fn several_values() {
        let ole = Reader::from_path("data/Thumbs.db").unwrap();
        assert_eq!(ole.sat.as_ref().unwrap().capacity(), 128usize);
        assert_eq!(ole.msat.as_ref().unwrap().len(), 1usize);
        assert_eq!(ole.ssat.as_ref().unwrap().capacity(), 512usize);
    }

    #[test]
    fn print_things() {
        use std::io::{Read, Write};
        let ole = Reader::from_path("data/sample.ppt").unwrap();
        println!(
            "STREAM SIZE: {}",
            *ole.minimum_standard_stream_size.as_ref().unwrap()
        );
        println!("MSAT: {:?}", ole.msat.as_ref().unwrap());
        println!("SAT: {:?}", ole.sat.as_ref().unwrap());
        println!("SSAT: {:?}", ole.ssat.as_ref().unwrap());
        println!("DSAT: {:?}", ole.dsat.as_ref().unwrap());
        let _ = std::fs::create_dir_all("data/streams");
        for entry in ole.iterate() {
            println!("{}", entry);
            if let Ok(mut slice) = ole.get_entry_slice(entry) {
                let mut buf = vec![0u8; slice.len()];
                let read_size = slice.read(&mut buf).unwrap();
                // Sanitize entry name for use as a filename (OLE names can
                // contain control characters like \x05 which are invalid on Windows)
                let safe_name: String = entry
                    .name()
                    .chars()
                    .map(|c| if c.is_alphanumeric() || c == '.' || c == '_' || c == '-' { c } else { '_' })
                    .collect();
                let mut file =
                    std::fs::File::create(format!("data/streams/{}.bin", safe_name)).unwrap();
                println!("Real len: {}", slice.real_len());
                file.write_all(&buf).unwrap();
                assert_eq!(read_size, slice.real_len());
                assert_eq!(read_size, slice.len());
            }
        }
    }
}
