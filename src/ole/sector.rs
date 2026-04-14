impl<'ole> super::ole::Reader<'ole> {
    pub(crate) fn read_sector(&self, sector_index: usize) -> Result<&[u8], super::error::Error> {
        let sector_size = self.sec_size.unwrap();
        let offset = sector_size
            .checked_mul(sector_index)
            .ok_or(super::error::Error::BadSizeValue("Sector offset overflow"))?;
        let max_size = offset
            .checked_add(sector_size)
            .ok_or(super::error::Error::BadSizeValue("Sector offset overflow"))?;

        let body = self
            .body
            .as_ref()
            .ok_or(super::error::Error::BadSizeValue("File is too short"))?;

        if body.len() >= max_size {
            Ok(&body[offset..offset + sector_size])
        } else {
            Err(super::error::Error::BadSizeValue("File is too short"))
        }
    }
}
