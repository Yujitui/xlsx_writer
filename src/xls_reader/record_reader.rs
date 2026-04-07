use byteorder::{LittleEndian, ReadBytesExt};
use std::io::{Read, Seek, SeekFrom};

// 辅助结构：用于自动处理跨记录读取
pub struct XlsRecordReader<'a, R: Read + Seek> {
    pub(crate) inner: &'a mut R,
    pub(crate) current_record_remaining: usize,
}

impl<'a, R: Read + Seek> XlsRecordReader<'a, R> {
    pub(crate) fn new(inner: &'a mut R, first_len: usize) -> Self {
        Self {
            inner,
            current_record_remaining: first_len,
        }
    }

    // 当当前 Record 读完时，检查下一个是不是 CONTINUE (0x003C)
    pub fn ensure_data(&mut self) -> std::io::Result<bool> {
        if self.current_record_remaining > 0 {
            return Ok(true);
        }

        // 尝试读取下一个 Record Header
        let next_type = match self.inner.read_u16::<LittleEndian>() {
            Ok(t) => t,
            Err(_) => return Ok(false),
        };

        if next_type == 0x003C { // CONTINUE 记录
            self.current_record_remaining = self.inner.read_u16::<LittleEndian>()? as usize;
            Ok(true)
        } else {
            // 不是 Continue，把头退回去，让外层循环处理
            self.inner.seek(SeekFrom::Current(-2))?;
            Ok(false)
        }
    }

    pub fn read_u8(&mut self) -> std::io::Result<u8> {
        if !self.ensure_data()? { return Err(std::io::Error::new(std::io::ErrorKind::UnexpectedEof, "No more data")); }
        let val = self.inner.read_u8()?;
        self.current_record_remaining -= 1;
        Ok(val)
    }

    pub fn read_exact(&mut self, buf: &mut [u8]) -> std::io::Result<()> {
        let mut read_so_far = 0;
        while read_so_far < buf.len() {
            if !self.ensure_data()? { return Err(std::io::Error::new(std::io::ErrorKind::UnexpectedEof, "Incomplete string")); }
            let to_read = std::cmp::min(buf.len() - read_so_far, self.current_record_remaining);
            self.inner.read_exact(&mut buf[read_so_far..read_so_far + to_read])?;
            self.current_record_remaining -= to_read;
            read_so_far += to_read;
        }
        Ok(())
    }
}