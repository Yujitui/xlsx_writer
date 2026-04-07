use thiserror::Error;
use std::io;

#[derive(Error, Debug)]
pub enum XlsError {
    #[error("IO error: {0}")]
    IoError(#[from] io::Error),

    #[error("Invalid file format: {0}")]
    InvalidFormat(String),

    #[error("String decoding failed: {0}")]
    Encoding(String),

    #[error("Cell index out of bounds: row={0}, col={1}")]
    OutOfBounds(usize, usize),

    #[error("Unexpected end of file")]
    UnexpectedEof,
}
