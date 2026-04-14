use thiserror::Error as ThisError;

/// Errors related to the process of parsing.
#[derive(ThisError, Debug)]
#[allow(clippy::enum_variant_names)]
pub enum Error {
    /// This happens when filesize is null, or to big to fit into an usize.
    #[error("Filesize is null or too big.")]
    BadFileSize,

    /// Classic std::io::Error.
    #[error(transparent)]
    IOError(#[from] std::io::Error),

    /// Something is not implemented yet ?
    #[error("Method not implemented yet")]
    NotImplementedYet,

    /// This is not a valid OLE file.
    #[error("Invalid OLE File")]
    InvalidOLEFile,

    /// Something has a bad size.
    #[error("{0}")]
    BadSizeValue(&'static str),

    /// MSAT is empty.
    #[error("MSAT is empty")]
    EmptyMasterSectorAllocationTable,

    /// Malformed SAT.
    #[error("Sector is not a sector used by the SAT.")]
    NotSectorUsedBySAT,

    /// Unknown node type.
    #[error("Unknown node type")]
    NodeTypeUnknown,

    /// Root storage has a bad size.
    #[error("Bad RootStorage size")]
    BadRootStorageSize,

    /// User query an empty entry
    #[error("Empty entry")]
    EmptyEntry,
}
