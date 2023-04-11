use std;
use std::io::Read;
use crate::ole::util::FromSlice;

impl<'ole> super::ole::Reader<'ole> {

  pub(crate) fn parse_header(&mut self) -> Result<(), super::error::Error> {
    // read the header
    let mut header: std::vec::Vec<u8>
        = vec![0u8; super::constants::HEADER_SIZE];
    self.read(&mut header)?;

    // initializes the return variable
    let result: Result<(), super::error::Error>;

    // Check file identifier
    if &super::constants::IDENTIFIER != &header[0..8] {
      result = Err(super::error::Error::InvalidOLEFile);
    } else {

      // UID
      self.uid = header[8..24].to_vec();

      // Revision number & version number
      let mut rv_number = usize::from_slice(&header[24..26]);
      self.revision_number = Some(rv_number as u16);
      rv_number = usize::from_slice(&header[26..28]);
      self.version_number = Some(rv_number as u16);

      // Check little-endianness; big endian not yet supported
      if &header[28..30] == &super::constants::BIG_ENDIAN_IDENTIFIER {
        result = Err(super::error::Error::NotImplementedYet);
      } else if
          &header[28..30] != &super::constants::LITTLE_ENDIAN_IDENTIFIER {
        result = Err(super::error::Error::InvalidOLEFile);
      } else {

        // Sector size
        let mut k = usize::from_slice(&header[30..32]);

        // if k >= 16, it means that the sector size equals 2 ^ k, which
        // is impossible.
        if k >= 16 {
          result =
            Err(super::error::Error::BadSizeValue("Overflow on sector
            size"));
        } else {
          self.sec_size = Some(2usize.pow(k as u32));


          // Short sector size
          k = usize::from_slice(&header[32..34]);

          // same for sector size
          if k >= 16 {
            result = Err(super::error::Error::BadSizeValue(
              "Overflow on short sector size"));
          } else {
            self.short_sec_size = Some(2usize.pow(k as u32));

            let sat: std::vec::Vec<u32>;


            // Total number of sectors used for the sector allocation table
            sat = std::vec::Vec::with_capacity(
              (*self.sec_size.as_ref().unwrap() / 4)
              *  usize::from_slice(&header[44..48]));

            // SecID of the first sector of directory stream
            let mut dsat: std::vec::Vec<u32> = std::vec::Vec::new();
            dsat.push(u32::from_slice(&header[48..52]));

            // Minimum size of a standard stream (bytes)
            self.minimum_standard_stream_size =
              Some(usize::from_slice(&header[56..60]));

            // standard says that this value has to be greater
            // or equals to 4096
            if *self.minimum_standard_stream_size.as_ref().unwrap()
                < 4096usize {
              result = Err(super::error::Error::InvalidOLEFile);
            } else {
              let mut ssat: std::vec::Vec<u32>;
              let mut msat: std::vec::Vec<u32>;

              // secID of the first sector of the SSAT & Total number
              // of sectors used for the short-sector allocation table
              ssat = std::vec::Vec::with_capacity(
                usize::from_slice(&header[64..68])
                * (*self.sec_size.as_ref().unwrap() / 4));
              ssat.push(u32::from_slice(&header[60..64]));

              // secID of first sector of the master sector allocation table
              // & Total number of sectors used for
              // the master sector allocation table
              msat = vec![super::constants::FREE_SECID_U32; 109];
              if &header[68..72] != &super::constants::END_OF_CHAIN_SECID {
                msat.resize(109usize + usize::from_slice(&header[72..76])
                  * (*self.sec_size.as_ref().unwrap() / 4),
                  super::constants::FREE_SECID_U32);
              }
              self.sat = Some(sat);
              self.msat = Some(msat);
              self.dsat = Some(dsat);
              self.ssat = Some(ssat);

              // now we build the MSAT
              self.build_master_sector_allocation_table(&header)?;
              result = Ok(())
            }
          }
        }
      }
    }

    result
  }


  /// Build the Master Sector Allocation Table (MSAT)
  fn build_master_sector_allocation_table(&mut self, header: &[u8])
      -> Result<(), super::error::Error> {

    // First, we build the master sector allocation table from the header
    let mut total_sec_id_read = self.read_sec_ids(&header[76 ..], 0);

    // Check if additional sectors are used for building the msat
    if total_sec_id_read == 109 {
      let sec_size = *self.sec_size.as_ref().unwrap();
      let mut sec_id = usize::from_slice(&header[68..72]);
      let mut buffer = vec![0u8; 0];
      let mut steps_since_last_resize = 0;


      while sec_id != super::constants::END_OF_CHAIN_SECID_U32 as usize {
        let relative_offset = sec_id * sec_size;

        // check if we need to read more data
        if buffer.len() < relative_offset + sec_size {
          let old_len = buffer.len();
          let new_len = relative_offset + sec_size;
          buffer.resize(new_len, 0xFFu8);
          self.read(&mut buffer[old_len..new_len])?;
          steps_since_last_resize = 0;
        }

        total_sec_id_read += self.read_sec_ids(&buffer[relative_offset
          .. relative_offset + sec_size - 4], total_sec_id_read);
        sec_id = usize::from_slice(&buffer[relative_offset + sec_size - 4
          .. relative_offset + sec_size]);

        steps_since_last_resize += 1;
        if steps_since_last_resize * sec_size > buffer.len() {
          // There is a loop in the MSAT chain
          return Err(super::error::Error::InvalidOLEFile);
        }
      }
        // save the buffer for later usage
        self.body = Some(buffer);
    }
    self.msat.as_mut().unwrap().resize(
      total_sec_id_read, super::constants::FREE_SECID_U32);

    // Now, we read the all file
    let mut buf: &mut std::vec::Vec<u8>;
    if !self.body.is_some() {
      self.body = Some(std::vec::Vec::new());
    }
    buf = self.body.as_mut().unwrap();

    self.buf_reader.as_mut().unwrap().read_to_end(&mut
      buf).map_err(super::error::Error::IOError)?;
    Ok(())
  }

  fn read_sec_ids(&mut self, buffer: &[u8], msat_offset: usize) -> usize {
    let mut i = 0usize;
    let mut offset = 0usize;
    let max_sec_ids = buffer.len() / 4;
    let msat = &mut self.msat.as_mut().unwrap()[msat_offset .. ];
    while i < max_sec_ids && &buffer[offset .. offset + 4]
      != &super::constants::FREE_SECID {
      msat[i] = u32::from_slice(&buffer[offset .. offset + 4]);
      offset += 4;
      i += 1;
    }

    i
  }
}
