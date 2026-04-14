pub(crate) trait FromSlice<T> {
    fn from_slice(buf: &[T]) -> Self;
}

impl FromSlice<u8> for usize {
    fn from_slice(buf: &[u8]) -> Self {
        let mut result = 0usize;
        for (p, &b) in buf.iter().enumerate() {
            result += (b as usize) * 256usize.pow(p as u32);
        }
        result
    }
}

impl FromSlice<u8> for u32 {
    fn from_slice(buf: &[u8]) -> Self {
        let mut result = 0u32;
        for (p, &b) in buf.iter().enumerate() {
            result += (b as u32) * 256u32.pow(p as u32);
        }
        result
    }
}

impl FromSlice<u8> for i32 {
    fn from_slice(buf: &[u8]) -> Self {
        let mut result = 0i32;
        for (p, &b) in buf.iter().enumerate() {
            result += (b as i32) * 256i32.pow(p as u32);
        }
        result
    }
}

impl FromSlice<u8> for u64 {
    fn from_slice(buf: &[u8]) -> Self {
        let mut result = 0u64;
        for (p, &b) in buf.iter().enumerate() {
            result += (b as u64) * 256u64.pow(p as u32);
        }
        result
    }
}
