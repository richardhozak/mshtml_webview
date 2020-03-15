use winapi::shared::guiddef::*;
use winapi::shared::minwindef::*;
use winapi::shared::ntdef::*;
use winapi::shared::windef::*;
use winapi::shared::winerror::{self, NOERROR};
use winapi::shared::wtypesbase::*;
use winapi::um::errhandlingapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::objidl::{IMoniker, FORMATETC, SNB};
use winapi::um::objidlbase::STATSTG;
use winapi::um::ole2::*;
use winapi::um::wingdi::LOGPALETTE;
use winapi::um::winuser::*;
// use winapi::um::objidlbase::IEnumUnknown;
use com::{co_class, com_interface, interfaces::iunknown::IUnknown, InterfacePtr};
use std::ffi::OsStr;
use std::os::raw::{c_char, c_void};
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

#[com_interface(00000112-0000-0000-C000-000000000046)]
pub trait IOleObject: IUnknown {
    unsafe fn set_client_site(&self, p_client_site: *mut c_void) -> HRESULT;
    unsafe fn get_client_site(&self, p_client_site: *mut *mut c_void) -> HRESULT;
    unsafe fn set_host_names(
        &self,
        sz_container_app: *const c_char,
        sz_container_obj: *const c_char,
    ) -> HRESULT;
    unsafe fn close(&self, dw_save_option: DWORD) -> HRESULT;
    unsafe fn set_moniker(&self, dw_which_moniker: DWORD, pmk: *mut c_void);
    unsafe fn get_moniker(
        &self,
        dw_assign: DWORD,
        dw_which_moniker: DWORD,
        ppmk: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn init_from_data(
        &self,
        p_data_object: *mut c_void,
        f_creation: BOOL,
        dw_reserved: DWORD,
    ) -> HRESULT;
    unsafe fn get_clipboard_data(
        &self,
        dw_reserved: DWORD,
        pp_data_object: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn do_verb(&self, i_verb: LONG, lpmsg: LPMSG, p_active_site: *mut c_void) -> HRESULT;
    unsafe fn enum_verbs(&self, pp_enum_ole_verb: *mut *mut c_void) -> HRESULT;
    unsafe fn update(&self) -> HRESULT;
    unsafe fn is_up_to_date(&self) -> HRESULT;
    unsafe fn get_user_class_id(&self, p_clsid: *mut CLSID) -> HRESULT;
    unsafe fn get_user_type(&self, dw_form_of_type: DWORD, psz_user_type: *mut LPOLESTR)
        -> HRESULT;
    unsafe fn set_extent(&self, dw_draw_aspect: DWORD, psizel: *mut SIZEL) -> HRESULT;
    unsafe fn get_extent(&self, dw_draw_aspect: DWORD, psizel: *mut SIZEL) -> HRESULT;
    unsafe fn advise(&self, p_advise_sink: *mut c_void, pdw_connection: *mut DWORD) -> HRESULT;
    unsafe fn unadvise(&self, dw_connection: DWORD) -> HRESULT;
    unsafe fn enum_advise(&self, ppenum_advise: *mut *mut c_void) -> HRESULT;
    unsafe fn get_misc_status(&self, dw_aspect: DWORD, pdw_status: *mut DWORD) -> HRESULT;
    unsafe fn set_color_scheme(&self, p_logpal: *mut LOGPALETTE) -> HRESULT;
}

#[com_interface(00000114-0000-0000-C000-000000000046)]
pub trait IOleWindow: IUnknown {
    unsafe fn get_window(&self, phwnd: *mut HWND) -> HRESULT;
    unsafe fn context_sensitive_help(&self, f_enter_mode: BOOL) -> HRESULT;
}

#[com_interface(00000113-0000-0000-C000-000000000046)]
pub trait IOleInPlaceObject: IOleWindow {
    unsafe fn in_place_deactivate(&self) -> HRESULT;
    unsafe fn ui_deactivate(&self) -> HRESULT;
    unsafe fn set_object_rects(&self, lprc_pos_rect: LPCRECT, lprc_clip_rect: LPCRECT) -> HRESULT;
    unsafe fn reactivate_and_undo(&self) -> HRESULT;
}

#[com_interface(00000119-0000-0000-C000-000000000046)]
pub trait IOleInPlaceSite: IOleWindow {
    unsafe fn can_in_place_activate(&self) -> HRESULT;
    unsafe fn on_in_place_activate(&self) -> HRESULT;
    unsafe fn on_ui_activate(&self) -> HRESULT;
    unsafe fn get_window_context(
        &self,
        pp_frame: *mut *mut c_void,
        pp_doc: *mut *mut c_void,
        lprc_pos_rect: LPRECT,
        lprc_clip_rect: LPRECT,
        lp_frame_info: *mut c_void,
    ) -> HRESULT;
    unsafe fn scroll(&self, scroll_extant: SIZE) -> HRESULT;
    unsafe fn on_ui_deactivate(&self, f_undoable: BOOL) -> HRESULT;
    unsafe fn on_in_place_deactivate(&self) -> HRESULT;
    unsafe fn discard_undo_state(&self) -> HRESULT;
    unsafe fn deactivate_and_undo(&self) -> HRESULT;
    unsafe fn on_pos_rect_change(&self, lprc_post_rect: LPRECT) -> HRESULT;
}

#[com_interface(0000000b-0000-0000-C000-000000000046)]
pub trait IStorage: IUnknown {
    unsafe fn create_stream(
        &self,
        pwcs_name: *const WCHAR,
        grf_mode: DWORD,
        reserved1: DWORD,
        reserved2: DWORD,
        ppstm: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn open_stream(
        &self,
        pwcs_name: *const WCHAR,
        reserved1: *mut c_void,
        grf_mode: DWORD,
        reserved2: DWORD,
        ppstm: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn create_storage(
        &self,
        pwcs_name: *const WCHAR,
        grf_mode: DWORD,
        reserved1: DWORD,
        reserved2: DWORD,
        ppstg: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn open_storage(
        &self,
        pwcs_name: *const WCHAR,
        pstg_priority: *mut c_void,
        grf_mode: DWORD,
        snb_exclude: SNB,
        reserved: DWORD,
        ppstg: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn copy_to(
        &self,
        ciid_exclude: DWORD,
        rgiid_exclude: *const IID,
        snb_exclude: SNB,
        pstg_dest: *mut c_void,
    ) -> HRESULT;
    unsafe fn move_element_to(
        &self,
        pwcs_name: *const WCHAR,
        pstg_dest: *mut c_void,
        pwcs_new_name: *const WCHAR,
        grf_flags: DWORD,
    ) -> HRESULT;
    unsafe fn commit(&self, grf_commit_flags: DWORD) -> HRESULT;
    unsafe fn revert(&self) -> HRESULT;
    unsafe fn enum_elements(
        &self,
        reserved1: DWORD,
        reserved2: *mut c_void,
        reserved3: DWORD,
        ppenum: *mut *mut c_void,
    ) -> HRESULT;
    unsafe fn destroy_element(&self, pwcs_name: *const WCHAR) -> HRESULT;
    unsafe fn rename_element(
        &self,
        pwcs_old_name: *const WCHAR,
        pwcs_new_name: *const WCHAR,
    ) -> HRESULT;
    unsafe fn set_element_times(
        &self,
        pwcs_name: *const WCHAR,
        pctime: *const FILETIME,
        patime: *const FILETIME,
        pmtime: *const FILETIME,
    ) -> HRESULT;
    unsafe fn set_class(&self, clsid: REFCLSID) -> HRESULT;
    unsafe fn set_state_bits(&self, grf_state_bits: DWORD, grf_mask: DWORD) -> HRESULT;
    unsafe fn stat(&self, pstatstg: *mut STATSTG, grf_stat_flag: DWORD) -> HRESULT;
}

#[co_class(implements(IOleClientSite, IOleInPlaceSite, IStorage))]
struct WebBrowser {
    //     ole_object: InterfacePtr<dyn IOleObject>,
//     ole_in_place_object: InterfacePtr<dyn IOleInPlaceObject>,
}

impl WebBrowser {
    fn new() -> Box<WebBrowser> {
        WebBrowser::allocate()
    }
}

impl IOleClientSite for WebBrowser {
    unsafe fn save_object(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn get_moniker(&self, _: u32, _: u32, _: *mut *mut std::ffi::c_void) -> i32 {
        unimplemented!()
    }
    unsafe fn get_container(&self, _: *mut *mut std::ffi::c_void) -> i32 {
        unimplemented!()
    }
    unsafe fn show_object(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn on_show_window(&self, _: i32) -> i32 {
        unimplemented!()
    }
    unsafe fn request_new_object_layout(&self) -> i32 {
        unimplemented!()
    }
}

impl IOleWindow for WebBrowser {
    unsafe fn get_window(&self, _: *mut *mut winapi::shared::windef::HWND__) -> i32 {
        unimplemented!()
    }
    unsafe fn context_sensitive_help(&self, _: i32) -> i32 {
        unimplemented!()
    }
}

impl IOleInPlaceSite for WebBrowser {
    unsafe fn can_in_place_activate(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn on_in_place_activate(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn on_ui_activate(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn get_window_context(
        &self,
        _: *mut *mut std::ffi::c_void,
        _: *mut *mut std::ffi::c_void,
        _: *mut winapi::shared::windef::RECT,
        _: *mut winapi::shared::windef::RECT,
        _: *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn scroll(&self, _: winapi::shared::windef::SIZE) -> i32 {
        unimplemented!()
    }
    unsafe fn on_ui_deactivate(&self, _: i32) -> i32 {
        unimplemented!()
    }
    unsafe fn on_in_place_deactivate(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn discard_undo_state(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn deactivate_and_undo(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn on_pos_rect_change(&self, _: *mut winapi::shared::windef::RECT) -> i32 {
        unimplemented!()
    }
}

impl IStorage for WebBrowser {
    unsafe fn create_stream(
        &self,
        _: *const u16,
        _: u32,
        _: u32,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn open_stream(
        &self,
        _: *const u16,
        _: *mut std::ffi::c_void,
        _: u32,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn create_storage(
        &self,
        _: *const u16,
        _: u32,
        _: u32,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn open_storage(
        &self,
        _: *const u16,
        _: *mut std::ffi::c_void,
        _: u32,
        _: *const *const u16,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn copy_to(
        &self,
        _: u32,
        _: *const winapi::shared::guiddef::GUID,
        _: *const *const u16,
        _: *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn move_element_to(
        &self,
        _: *const u16,
        _: *mut std::ffi::c_void,
        _: *const u16,
        _: u32,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn commit(&self, _: u32) -> i32 {
        unimplemented!()
    }
    unsafe fn revert(&self) -> i32 {
        unimplemented!()
    }
    unsafe fn enum_elements(
        &self,
        _: u32,
        _: *mut std::ffi::c_void,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn destroy_element(&self, _: *const u16) -> i32 {
        unimplemented!()
    }
    unsafe fn rename_element(&self, _: *const u16, _: *const u16) -> i32 {
        unimplemented!()
    }
    unsafe fn set_element_times(
        &self,
        _: *const u16,
        _: *const winapi::shared::minwindef::FILETIME,
        _: *const winapi::shared::minwindef::FILETIME,
        _: *const winapi::shared::minwindef::FILETIME,
    ) -> i32 {
        unimplemented!()
    }
    unsafe fn set_class(&self, _: *const winapi::shared::guiddef::GUID) -> i32 {
        unimplemented!()
    }
    unsafe fn set_state_bits(&self, _: u32, _: u32) -> i32 {
        unimplemented!()
    }
    unsafe fn stat(&self, _: *mut winapi::um::objidlbase::STATSTG, _: u32) -> i32 {
        unimplemented!()
    }
}

// #[com_interface(0000011b-0000-0000-C000-000000000046)]
// pub trait IOleContainer: IUnknown {
//     unsafe fn enum_objects(&self, grf_flags: DWORD, ppenum: *mut *mut IEnumUnknown) -> HRESULT;
//     unsafe fn lock_container(&self, f_lock: BOOL) -> HRESULT;
// }

fn main() {
    unsafe {
        let result = OleInitialize(ptr::null_mut());
        if result != winerror::S_OK && result != winerror::S_FALSE {
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
        }
        WM_SIZE => {
            println!("size");
            1
        }
        WM_DESTROY => {
            unsafe {
                ExitProcess(0);
            }
            1
        }
        _ => unsafe { DefWindowProcW(hwnd, message, wparam, lparam) },
    }
}

fn to_wstring(s: &str) -> Vec<u16> {
    OsStr::new(s)
        .encode_wide()
        .chain(Some(0).into_iter())
        .collect()
}
