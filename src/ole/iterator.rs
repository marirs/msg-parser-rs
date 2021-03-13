/// Iterator for entries inside an OLE file.
pub struct OLEIterator<'a> {
  ole: &'a super::ole::Reader<'a>,
  curr: usize
}

impl<'a> OLEIterator<'a> {

  pub(crate) fn new(ole: &'a super::ole::Reader) -> OLEIterator<'a> {
    OLEIterator {
      ole: ole,
      curr: 0
    }
  }
}

impl<'a> Iterator for OLEIterator<'a> {
  type Item = &'a super::entry::Entry;

  fn next(&mut self) -> Option<&'a super::entry::Entry> {
    let entries = self.ole.entries.as_ref().unwrap();
    if self.curr < entries.len() {
      self.curr += 1;
      Some(&entries[self.curr - 1])
    } else {
      None
    }
  }
}
