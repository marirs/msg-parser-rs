use std;
use crate::ole::util::FromSlice;

impl<'ole> super::ole::Reader<'ole> {

  pub(crate) fn build_sat(&mut self)
    -> Result<(), super::error::Error> {
    let sector_size = self.sec_size.unwrap();
    let result: Result<(), super::error::Error>;
    let mut sec_ids = vec![super::constants::FREE_SECID_U32;
        sector_size / 4];
    if self.msat.as_ref().unwrap().len() == 0 {
      result = Err(super::error::Error::EmptyMasterSectorAllocationTable);
    } else {
      for i in 0 .. self.msat.as_ref().unwrap().len() {
        let sector_index = self.msat.as_ref().unwrap()[i];
        self.read_sat_sector(sector_index as usize, &mut sec_ids)?;
        self.sat.as_mut().unwrap().extend_from_slice(&sec_ids);
      }
      self.build_ssat()?;
      self.build_dsat()?;
      result = Ok(());
    }
    result
  }

  pub(crate) fn read_sat_sector(&mut self, sector_index: usize,
      sec_ids: &mut std::vec::Vec<u32> ) -> Result<(), super::error::Error> {
    let sector = self.read_sector(sector_index)?;
    for i in 0 .. sec_ids.capacity() {
      sec_ids[i] = u32::from_slice(&sector[ i * 4 .. i * 4 + 4]);
    }

    Ok(())
  }

  pub(crate) fn build_chain_from_sat(&mut self, start: u32)
        -> std::vec::Vec<u32> {
    let mut chain = std::vec::Vec::new();
    let mut sector_index = start;
    let sat = self.sat.as_mut().unwrap();
    while sector_index != super::constants::END_OF_CHAIN_SECID_U32 {
      chain.push(sector_index);
      sector_index = sat[sector_index as usize];
    }

    chain
  }

  pub(crate) fn build_chain_from_ssat(&mut self, start: u32)
        -> std::vec::Vec<u32> {
    let mut chain = std::vec::Vec::new();
    let mut sector_index = start;
    let sat = self.ssat.as_mut().unwrap();
    while sector_index != super::constants::END_OF_CHAIN_SECID_U32
        && sector_index != super::constants::FREE_SECID_U32 {
      chain.push(sector_index);

      sector_index = sat[sector_index as usize];
    }

    chain
  }

  pub(crate) fn build_ssat(&mut self) -> Result<(), super::error::Error> {
    let mut sec_ids = vec![super::constants::FREE_SECID_U32;
        self.sec_size.as_ref().unwrap() / 4];

    let sector_index = self.ssat.as_mut().unwrap().remove(0);
    let chain = self.build_chain_from_sat(sector_index);

    for sector_index in chain {
      self.read_sat_sector(sector_index as usize, &mut sec_ids)?;
      self.ssat.as_mut().unwrap().extend_from_slice(&sec_ids);
    }
    Ok(())
  }

  pub(crate) fn build_dsat(&mut self) -> Result<(), super::error::Error> {

    let sector_index = self.dsat.as_mut().unwrap().remove(0);
    let chain = self.build_chain_from_sat(sector_index);

    for sector_index in chain {
      self.dsat.as_mut().unwrap().push(sector_index);
    }

    Ok(())
  }
}
