mod constants;
mod decode;
mod storage;
mod stream;

mod error;
pub use error::{DataTypeError, Error};

mod outlook;
pub use outlook::{Attachment, Outlook, Person, TransportHeaders};
