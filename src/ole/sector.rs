impl<'ole> super::ole::Reader<'ole> {
  pub(crate) fn read_sector(&self, sector_index: usize)
    -> Result<&[u8], super::error::Error> {
    let result: Result<&[u8], super::error::Error>;
    let sector_size = self.sec_size.unwrap();
    let offset = sector_size * sector_index;
    let max_size = offset + sector_size;

    let body_size: usize;
    if self.body.is_some() {
      body_size = self.body.as_ref().unwrap().len();
    } else {
      body_size = 0;
    }

    // Check if the sector has already been read
    let sector : &[u8];
    if body_size >= max_size {
      let body = self.body.as_ref().unwrap();
      sector = &body[offset .. offset + sector_size];
      result = Ok(sector);
    } else {
      result = Err(super::error::Error::BadSizeValue("File is too short"));
    }

    result
  }
}
