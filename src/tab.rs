use std::io;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::Win32::Foundation::HANDLE;

use crate::pty::{Pty, ShellType};
use crate::term::TermWrapper;

pub type TabId = u32;

static NEXT_ID: AtomicU32 = AtomicU32::new(1);

fn next_tab_id() -> TabId {
    NEXT_ID.fetch_add(1, Ordering::Relaxed)
}

pub struct Tab {
    pub id: TabId,
    pub pty: Pty,
    pub term: TermWrapper,
    pub shell_type: ShellType,
    pub title: String,
}

impl Tab {
    pub fn new(cols: usize, rows: usize, shell: ShellType) -> io::Result<Self> {
        let pty = Pty::new(cols as u16, rows as u16, shell)?;
        let term = TermWrapper::new(cols, rows);
        Ok(Self {
            id: next_tab_id(),
            pty,
            term,
            shell_type: shell,
            title: shell.label().to_string(),
        })
    }

    pub fn dup_output_read(&self) -> io::Result<HANDLE> {
        self.pty.dup_output_read()
    }

    pub fn dup_process_handle(&self) -> io::Result<HANDLE> {
        self.pty.dup_process_handle()
    }

    pub fn process_pty_output(&mut self, data: &[u8]) {
        self.term.process(data);
    }

    pub fn write_pty(&self, data: &[u8]) -> io::Result<usize> {
        self.pty.write(data)
    }

    pub fn resize(&mut self, cols: usize, rows: usize) {
        self.term.resize(cols, rows);
        let _ = self.pty.resize(cols as u16, rows as u16);
    }
}
