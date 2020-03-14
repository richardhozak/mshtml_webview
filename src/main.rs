use winapi::shared::guiddef::*;
use winapi::shared::minwindef::*;
use winapi::shared::windef::HWND;
use winapi::shared::winerror::*;
use winapi::um::errhandlingapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::objidl::{IMoniker, FORMATETC};
use winapi::um::ole2::*;
use winapi::um::winuser::*;
// use winapi::um::objidlbase::IEnumUnknown;
use com::{com_interface, interfaces::iunknown::IUnknown};
use std::ffi::OsStr;
use std::os::raw::c_void;
use std::os::windows::ffi::OsStrExt;
use std::ptr;

type LPFORMATETC = *mut FORMATETC;

extern "stdcall" {
    fn OleCreate(
        rclsid: REFCLSID,
        riid: REFIID,
        renderopt: DWORD,
        pFormatEtc: LPFORMATETC,
        p_client_size: *mut c_void,
        p_str: *mut c_void,
        ppv_obj: *mut *mut c_void,
    );

    fn ExitProcess(exit_code: UINT);
}

extern "system" {
    fn OleUninitialize();
}

#[com_interface(00000118-0000-0000-C000-000000000046)]
pub trait IOleClientSite: IUnknown {
    unsafe fn save_object(&self) -> HRESULT;
    unsafe fn get_moniker(
        &self,
        dw_assign: DWORD,
        dw_which_moniker: DWORD,
        ppmk: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn get_container(&self, pp_container: *mut *mut c_void) -> HRESULT;
    unsafe fn show_object(&self) -> HRESULT;
    unsafe fn on_show_window(&self, show: BOOL) -> HRESULT;
    unsafe fn request_new_object_layout(&self) -> HRESULT;
}

// #[com_interface(0000011b-0000-0000-C000-000000000046)]
// pub trait IOleContainer: IUnknown {
//     unsafe fn enum_objects(&self, grf_flags: DWORD, ppenum: *mut *mut IEnumUnknown) -> HRESULT;
//     unsafe fn lock_container(&self, f_lock: BOOL) -> HRESULT;
// }

fn main() {
    unsafe {
        let result = OleInitialize(ptr::null_mut());
        if result != S_OK && result != S_FALSE {
            panic!("could not initialize ole");
        }

        let h_instance = GetModuleHandleA(ptr::null_mut());
        if h_instance.is_null() {
            panic!("could not retrieve module handle");
        }

        let class_name = to_wstring("webview");
        let class = WNDCLASSW {
            style: 0,
            lpfnWndProc: Some(wndproc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: h_instance,
            hIcon: ptr::null_mut(),
            hCursor: LoadCursorW(ptr::null_mut(), IDC_ARROW),
            hbrBackground: ptr::null_mut(),
            lpszMenuName: ptr::null(),
            lpszClassName: class_name.as_ptr(),
        };

        if RegisterClassW(&class) == 0 {
            // ignore the "Class already exists" error for multiple windows
            if GetLastError() as u32 != 1410 {
                OleUninitialize();
                panic!("could not register window class {}", GetLastError() as u32);
            }
        }

        let title = to_wstring("mshtml_webview");
        let handle = CreateWindowExW(
            0,
            class_name.as_ptr(),
            title.as_ptr(),
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            HWND_DESKTOP,
            ptr::null_mut(),
            h_instance,
            ptr::null_mut(),
        );

        ShowWindow(handle, SW_SHOWDEFAULT);

        let mut message: MSG = Default::default();
        while GetMessageW(&mut message, ptr::null_mut(), 0, 0) > 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }
}

extern "system" fn wndproc(hwnd: HWND, message: UINT, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match message {
        WM_CREATE => {
            println!("created");
            1
        },
        WM_SIZE => {
            println!("size");
            1
        },
        WM_DESTROY => {
            unsafe { ExitProcess(0); }
            1
        },
        _ => unsafe { DefWindowProcW(hwnd, message, wparam, lparam) }
    }
}

fn to_wstring(s: &str) -> Vec<u16> {
    OsStr::new(s)
        .encode_wide()
        .chain(Some(0).into_iter())
        .collect()
}
