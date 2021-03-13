use std;

/// An OLE file reader.
///
/// The parsing method follows the same method described here:
/// <http://www.openoffice.org/sc/compdocfileformat.pdf>
///
/// # Basic Example
///
/// ```ignore
/// use crate::ole::Reader;
///
/// let mut reader =
///       Reader::from_path("assets/Thumbs.db").unwrap();
///
/// println!("This OLE file stores the following entries: ");
/// for entry in reader.iterate() {
///   println!("{}", entry);
/// }
/// ```

pub struct Reader<'ole> {

  /// Buffer for reading from the source.
  pub(crate) buf_reader: Option<std::io::BufReader<Box<std::io::Read + 'ole>>>,

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
  pub(crate) root_entry: Option<u32>
}

impl<'ole> Reader<'ole> {

  /// Constructs a new `Reader`.
  ///
  /// # Examples
  ///
  /// ```ignore
  /// use ole;
  /// let mut my_resume = std::fs::File::open("assets/Thumbs.db").unwrap();
  /// let mut parser = ole::Reader::new(my_resume).unwrap();
  /// ```
  pub fn new<T: 'ole>(readable: T)
        -> std::result::Result<Reader<'ole>, super::error::Error>
    where T: std::io::Read {
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
      root_entry: None
    };
    t.parse_header()?;
    t.build_sat()?;
    t.build_directory_entries()?;
    Ok(t)
  }


  /// Constructs a new `Reader` from a file.
  ///
  /// # Examples
  ///
  /// ```ignore
  /// use ole;
  /// let mut parser = ole::Reader::from_path("assets/Thumbs.db").unwrap();
  /// ```
  pub fn from_path(path: &str) -> Result<Reader, super::error::Error> {
    let f = std::fs::File::open(path).map_err(super::error::Error::IOError)?;
    Reader::new(f)
  }


  /// Returns an iterator for directory entries of the OLE file.
  ///
  /// # Examples
  ///
  /// ```ignore
  /// use ole;
  /// let mut parser = ole::Reader::from_path("assets/Thumbs.db").unwrap();
  ///
  /// for entry in parser.iterate() {
  ///   println!("Entry {}", entry.name());
  /// }
  /// ```
  pub fn iterate(&self) -> super::iterator::OLEIterator {
    super::iterator::OLEIterator::new(self)
  }

  /// Read some bytes from the source.
  pub(crate) fn read(&mut self, buf: &mut [u8])
        -> Result<usize, super::error::Error> {
    use std::io::Read;
    self.buf_reader.as_mut().unwrap().read_exact(buf)
        .map_err(super::error::Error::IOError)?;
    Ok(buf.len())
  }

}


#[cfg(test)]
mod tests {
  use std;
  use super::Reader;
  use std::error::Error as e;
  use super::super::error::Error;

  #[test]
  fn instance_nok() {
    let path = "Thumbs.db";
    let o : Result<Reader, Error> = Reader::from_path(path);
    assert_eq!(o.is_ok(), false);
    let e = o.err().unwrap();
    println!("NOK: {}", e.description());
  }

  #[test]
  fn instance_ok() {
    let path = "data/Thumbs.db";
    let o: Result<Reader, Error> = Reader::from_path(path);
    assert_eq!(o.is_ok(), true);
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
    assert_eq!(ole.is_ok(), false);
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
    assert_eq!(ole.is_ok(), false);
    println!("BAD ENDIANNESS: {}", ole.err().unwrap());
  }

  #[test]
  fn uid() {
    let ole = Reader::from_path("data/Thumbs.db");
    assert_eq!(ole.is_ok(), true);
    let ole = ole.unwrap();
    assert_eq!(&[0x0u8; 16] == &ole.uid[..], true);
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
    assert_eq!(ole.is_ok(), false);
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
    println!("STREAM SIZE: {}", *ole.minimum_standard_stream_size.as_ref().unwrap());
    println!("MSAT: {:?}", ole.msat.as_ref().unwrap());
    println!("SAT: {:?}", ole.sat.as_ref().unwrap());
    println!("SSAT: {:?}", ole.ssat.as_ref().unwrap());
    println!("DSAT: {:?}", ole.dsat.as_ref().unwrap());
    for entry in ole.iterate() {
      println!("{}", entry);
      if let Ok(mut slice) = ole.get_entry_slice(entry) {
        let mut buf = vec![0u8; slice.len()];
        let read_size = slice.read(&mut buf).unwrap();
        let mut file = std::fs::File::create(format!("data/streams/{}.bin", entry.name())).unwrap();
        println!("Real len: {}", slice.real_len());
        file.write_all(&buf).unwrap();
        assert_eq!(read_size, slice.real_len());
        assert_eq!(read_size, slice.len());
      }
    }
  }
}
