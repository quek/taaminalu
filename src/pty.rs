use std::io;
use std::mem;
use std::ptr;

use windows::Win32::Foundation::{CloseHandle, DuplicateHandle, DUPLICATE_SAME_ACCESS, HANDLE};
use windows::Win32::System::Threading::GetCurrentProcess;
use windows::Win32::Security::SECURITY_ATTRIBUTES;
use windows::Win32::System::Console::{
    ClosePseudoConsole, CreatePseudoConsole, ResizePseudoConsole, COORD, HPCON,
};
use windows::Win32::System::Pipes::CreatePipe;
use windows::Win32::System::Threading::{
    CreateProcessW, DeleteProcThreadAttributeList,
    InitializeProcThreadAttributeList, UpdateProcThreadAttribute, EXTENDED_STARTUPINFO_PRESENT,
    LPPROC_THREAD_ATTRIBUTE_LIST, PROCESS_INFORMATION, PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,
    STARTF_USESTDHANDLES, STARTUPINFOEXW, STARTUPINFOW,
};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ShellType {
    Wsl,
    Cmd,
    PowerShell,
}

impl ShellType {
    pub fn label(&self) -> &'static str {
        match self {
            ShellType::Wsl => "WSL",
            ShellType::Cmd => "CMD",
            ShellType::PowerShell => "PowerShell",
        }
    }

    fn command(&self) -> &'static str {
        match self {
            ShellType::Wsl => "wsl.exe\0",
            ShellType::Cmd => "cmd.exe\0",
            ShellType::PowerShell => "powershell.exe\0",
        }
    }
}

pub struct Pty {
    hpc: HPCON,
    input_write: HANDLE,
    output_read: HANDLE,
    process: HANDLE,
    thread: HANDLE,
}

// HANDLE は Send safe（カーネルオブジェクト）
unsafe impl Send for Pty {}
unsafe impl Sync for Pty {}

/// ハンドルを複製して返す
fn dup_handle(src: HANDLE) -> io::Result<HANDLE> {
    let mut dup = HANDLE::default();
    unsafe {
        let process = GetCurrentProcess();
        DuplicateHandle(process, src, process, &mut dup, 0, false, DUPLICATE_SAME_ACCESS)
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
    }
    Ok(dup)
}

impl Pty {
    pub fn new(cols: u16, rows: u16, shell: ShellType) -> io::Result<Self> {
        unsafe { Self::create(cols, rows, shell) }
    }

    unsafe fn create(cols: u16, rows: u16, shell: ShellType) -> io::Result<Self> {
        let size = COORD {
            X: cols as i16,
            Y: rows as i16,
        };

        // ConPTY 入力パイプ: input_read → ConPTY が読む, input_write → 我々が書く
        let mut input_read = HANDLE::default();
        let mut input_write = HANDLE::default();
        // ConPTY 出力パイプ: output_read → 我々が読む, output_write → ConPTY が書く
        let mut output_read = HANDLE::default();
        let mut output_write = HANDLE::default();

        let sa = SECURITY_ATTRIBUTES {
            nLength: mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
            bInheritHandle: true.into(),
            lpSecurityDescriptor: ptr::null_mut(),
        };

        unsafe {
            CreatePipe(&mut input_read, &mut input_write, Some(&sa), 0)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
            CreatePipe(&mut output_read, &mut output_write, Some(&sa), 0)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
        }

        // ConPTY 作成
        let hpc = unsafe { CreatePseudoConsole(size, input_read, output_write, 0) }
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;

        // ConPTY に渡したパイプ端はもう不要
        unsafe {
            let _ = CloseHandle(input_read);
            let _ = CloseHandle(output_write);
        }

        // プロセス属性リスト作成
        let mut attr_size: usize = 0;
        let _ = unsafe {
            InitializeProcThreadAttributeList(None, 1, None, &mut attr_size)
        };

        let mut attr_buf = vec![0u8; attr_size];
        let attr_list = LPPROC_THREAD_ATTRIBUTE_LIST(attr_buf.as_mut_ptr() as *mut _);

        unsafe {
            InitializeProcThreadAttributeList(Some(attr_list), 1, None, &mut attr_size)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;

            UpdateProcThreadAttribute(
                attr_list,
                0,
                PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE as usize,
                Some(hpc.0 as *const _),
                mem::size_of::<HPCON>(),
                None,
                None,
            )
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
        }

        // STARTF_USESTDHANDLES + null ハンドルで親のコンソール継承を防ぐ
        let si = STARTUPINFOEXW {
            StartupInfo: STARTUPINFOW {
                cb: mem::size_of::<STARTUPINFOEXW>() as u32,
                dwFlags: STARTF_USESTDHANDLES,
                ..Default::default()
            },
            lpAttributeList: attr_list,
        };

        let mut pi = PROCESS_INFORMATION::default();
        let mut cmd: Vec<u16> = shell.command().encode_utf16().collect();

        unsafe {
            CreateProcessW(
                None,
                Some(windows::core::PWSTR(cmd.as_mut_ptr())),
                None,
                None,
                false,
                EXTENDED_STARTUPINFO_PRESENT,
                None,
                None,
                &si.StartupInfo,
                &mut pi,
            )
            .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;

            DeleteProcThreadAttributeList(attr_list);
        }

        Ok(Pty {
            hpc,
            input_write,
            output_read,
            process: pi.hProcess,
            thread: pi.hThread,
        })
    }

    /// PTY に書き込み
    pub fn write(&self, data: &[u8]) -> io::Result<usize> {
        use windows::Win32::Storage::FileSystem::WriteFile;
        let mut bytes_written = 0u32;
        unsafe {
            WriteFile(self.input_write, Some(data), Some(&mut bytes_written), None)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
        }
        Ok(bytes_written as usize)
    }

    /// プロセスハンドルを複製して返す（プロセス監視スレッド用）
    pub fn dup_process_handle(&self) -> io::Result<HANDLE> {
        dup_handle(self.process)
    }

    /// 読み取りハンドルを複製して返す（読み取りスレッド用）
    pub fn dup_output_read(&self) -> io::Result<HANDLE> {
        dup_handle(self.output_read)
    }

    /// ターミナルサイズ変更
    pub fn resize(&self, cols: u16, rows: u16) -> io::Result<()> {
        let size = COORD {
            X: cols as i16,
            Y: rows as i16,
        };
        unsafe {
            ResizePseudoConsole(self.hpc, size)
                .map_err(|e| io::Error::new(io::ErrorKind::Other, e))?;
        }
        Ok(())
    }
}

impl Drop for Pty {
    fn drop(&mut self) {
        unsafe {
            ClosePseudoConsole(self.hpc);
            let _ = CloseHandle(self.process);
            let _ = CloseHandle(self.thread);
            let _ = CloseHandle(self.input_write);
            let _ = CloseHandle(self.output_read);
        }
    }
}
