use winapi::shared::minwindef::*;
use winapi::shared::ntdef::*;
use winapi::shared::windef::*;
use winapi::shared::winerror::{self, FAILED, S_OK};
use winapi::um::errhandlingapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::objidl::FORMATETC;
use winapi::um::ole2::*;
use winapi::um::winuser::*;
use com::{co_class, interfaces::iunknown::IUnknown, ComPtr};
use libc::c_void;
use std::ffi::{OsStr, OsString};
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::ptr;

mod interface;
mod interface_impl;

use interface::*;

// "8856F961-340A-11D0-A96B-00C04FD705A2"
#[allow(non_upper_case_globals)]
const CLSID_WebBrowser: com::sys::IID = com::sys::IID {
    data1: 0x8856F961,
    data2: 0x340A,
    data3: 0x11D0,
    data4: [0xA9, 0x6B, 0x00, 0xC0, 0x4F, 0xD7, 0x05, 0xA2],
};

type LPFORMATETC = *mut FORMATETC;

extern "stdcall" {
    fn OleCreate(
        rclsid: *const com::sys::IID,
        riid: *const com::sys::IID,
        renderopt: DWORD,
        pFormatEtc: LPFORMATETC,
        p_client_size: *mut c_void,
        p_str: *mut c_void,
        ppv_obj: *mut *mut c_void,
    ) -> HRESULT;

    fn OleSetContainedObject(p_unknown: *mut c_void, f_contained: BOOL) -> HRESULT;

    fn ExitProcess(exit_code: UINT);
}

extern "system" {
    fn OleUninitialize();
}

#[co_class(implements(IOleClientSite, IOleInPlaceSite, IStorage))]
struct WebBrowser {
    inner: Option<WebBrowserInner>,
}

struct WebBrowserInner {
    hwnd_parent: HWND,
    rect: RECT,
    ole_in_place_object: ComPtr<dyn IOleInPlaceObject>,
    web_browser: ComPtr<dyn IWebBrowser>,
}

impl WebBrowser {
    /// A safe version of `QueryInterface`. If the backing CoClass implements the
    /// interface `I` then a `Some` containing an `ComRc` pointing to that
    /// interface will be returned otherwise `None` will be returned.
    fn get_interface<I: com::ComInterface + ?Sized>(&self) -> Option<ComPtr<I>> {
        let mut ppv = std::ptr::null_mut::<c_void>();
        let hr = unsafe { self.query_interface(&I::IID as *const com::sys::IID, &mut ppv) };
        if FAILED(hr) {
            assert!(
                hr == com::sys::E_NOINTERFACE || hr == com::sys::E_POINTER,
                "QueryInterface returned non-standard error"
            );
            return None;
        }
        assert!(!ppv.is_null(), "The pointer to the interface returned from a successful call to QueryInterface was null");
        Some(unsafe { ComPtr::new(ppv as *mut *mut _) })
    }

    fn new() -> Box<WebBrowser> {
        WebBrowser::allocate(None)
    }

    fn set_rect(&self, mut rect: RECT) {
        if self.inner.is_none() {
            return;
        }

        rect.top = 45;
        unsafe {
            self.inner
                .as_ref()
                .unwrap()
                .ole_in_place_object
                .set_object_rects(&rect, &rect);
        }
    }

    fn navigate(&self, url: &str) {
        let mut wstring = to_wstring(url);
        unsafe {
            self.inner.as_ref().unwrap().web_browser.navigate(
                wstring.as_mut_ptr(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
            );
        }
    }

    fn prev(&self) {
        unsafe {
            self.inner.as_ref().unwrap().web_browser.go_back();
        }
    }

    fn next(&self) {
        unsafe {
            self.inner.as_ref().unwrap().web_browser.go_forward();
        }
    }

    fn refresh(&self) {
        unsafe {
            self.inner.as_ref().unwrap().web_browser.refresh();
        }
    }

    fn initialize(&mut self, h_wnd: HWND, rect: RECT) {
        unsafe {
            let iole_client_site = self
                .get_interface::<dyn IOleClientSite>()
                .expect("iole_client_site query failed");

            let istorage = self
                .get_interface::<dyn IStorage>()
                .expect("istorage query failed");

            let mut ioleobject_ptr = ptr::null_mut::<c_void>();
            let hresult = OleCreate(
                &CLSID_WebBrowser,
                &<dyn IOleObject as com::ComInterface>::IID,
                1,
                ptr::null_mut(),
                iole_client_site.as_raw() as _,
                istorage.as_raw() as _,
                &mut ioleobject_ptr,
            );

            if FAILED(hresult) {
                panic!("cannot create WebBrowser ole object");
            }

            let ioleobject = ComPtr::<dyn IOleObject>::new(ioleobject_ptr as *mut *mut _);
            let hresult = OleSetContainedObject(ioleobject.as_raw() as _, 1);

            if FAILED(hresult) {
                panic!("OleSetContainedObject() failed");
            }

            let ole_in_place_object = ioleobject
                .get_interface::<dyn IOleInPlaceObject>()
                .expect("cannot query ole_in_place_object");

            ole_in_place_object.set_object_rects(&rect, &rect);
            let mut hwnd_control: HWND = ptr::null_mut();
            ole_in_place_object.get_window(&mut hwnd_control);
            assert!(!hwnd_control.is_null(), "in place object hwnd is null");

            let web_browser = ioleobject
                .get_interface::<dyn IWebBrowser>()
                .expect("get interface IWebBrowser failed");

            self.inner = Some(WebBrowserInner {
                hwnd_parent: h_wnd,
                rect,
                ole_in_place_object,
                web_browser,
            });

            let hresult = ioleobject.do_verb(
                -5,
                ptr::null_mut(),
                iole_client_site.as_raw() as _,
                -1,
                h_wnd,
                &rect,
            );

            if FAILED(hresult) {
                panic!("ioleobject.do_verb() failed");
            }
        }
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
        if result != S_OK && result != winerror::S_FALSE {
            panic!("could not initialize ole");
        }

        let h_instance = GetModuleHandleW(ptr::null_mut());
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
            hbrBackground: COLOR_WINDOW as _,
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
        let h_wnd = CreateWindowExW(
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

        let mut wb = WebBrowser::new();
        wb.initialize(
            h_wnd,
            RECT {
                left: 0,
                right: 300,
                top: 45,
                bottom: 300,
            },
        );
        wb.navigate("http://google.com");

        let wb_ptr = Box::into_raw(wb);

        SetWindowLongPtrW(h_wnd, GWLP_USERDATA, std::mem::transmute(wb_ptr));
        ShowWindow(h_wnd, SW_SHOWDEFAULT);

        let mut message: MSG = Default::default();
        while GetMessageW(&mut message, ptr::null_mut(), 0, 0) > 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }

        let _ = Box::from_raw(wb_ptr);
    }
}

const BTN_BACK: WORD = 1;
const BTN_NEXT: WORD = 2;
const BTN_REFRESH: WORD = 3;
const BTN_GO: WORD = 4;

unsafe extern "system" fn wndproc(
    hwnd: HWND,
    message: UINT,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    static mut EDIT_HWND: HWND = ptr::null_mut();

    match message {
        WM_CREATE => {
            let h_instance = GetModuleHandleW(ptr::null_mut());
            if h_instance.is_null() {
                panic!("could not retrieve module handle");
            }

            CreateWindowExW(
                0,
                to_wstring("BUTTON").as_ptr(),
                to_wstring("<<< Back").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                5,
                5,
                80,
                30,
                hwnd,
                BTN_BACK as _,
                h_instance,
                ptr::null_mut(),
            );

            CreateWindowExW(
                0,
                to_wstring("BUTTON").as_ptr(),
                to_wstring("Next >>>").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                90,
                5,
                80,
                30,
                hwnd,
                BTN_NEXT as _,
                h_instance,
                ptr::null_mut(),
            );

            CreateWindowExW(
                0,
                to_wstring("BUTTON").as_ptr(),
                to_wstring("Refresh").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                175,
                5,
                80,
                30,
                hwnd,
                BTN_REFRESH as _,
                h_instance,
                ptr::null_mut(),
            );

            EDIT_HWND = CreateWindowExW(
                0,
                to_wstring("EDIT").as_ptr(),
                to_wstring("http://google.com/").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_BORDER,
                260,
                10,
                200,
                20,
                hwnd,
                ptr::null_mut(),
                h_instance,
                lparam as _,
            );

            CreateWindowExW(
                0,
                to_wstring("BUTTON").as_ptr(),
                to_wstring("Go").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                465,
                5,
                80,
                30,
                hwnd,
                BTN_GO as _,
                h_instance,
                ptr::null_mut(),
            );

            1
        }
        WM_COMMAND => {
            let wb_ptr: *mut WebBrowser =
                std::mem::transmute(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if wb_ptr.is_null() {
                return 1;
            }
            let cmd = LOWORD(wparam as _);
            match cmd {
                BTN_BACK => (*wb_ptr).prev(),
                BTN_NEXT => (*wb_ptr).next(),
                BTN_REFRESH => (*wb_ptr).refresh(),
                BTN_GO => {
                    let mut buf: [u16; 4096] = [0; 4096];

                    let len = GetWindowTextW(EDIT_HWND, buf.as_mut_ptr(), buf.len() as _);
                    let len = len as usize;

                    if len == 0 {
                        return 1;
                    }

                    let input = OsString::from_wide(&buf[..len + 1]);
                    (*wb_ptr).navigate(&input.to_string_lossy());
                }
                _ => {}
            }

            1
        }
        WM_SIZE => {
            let wb_ptr: *mut WebBrowser =
                std::mem::transmute(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if wb_ptr.is_null() {
                return 1;
            }
            let mut rect: RECT = Default::default();
            GetClientRect(hwnd, &mut rect);
            (*wb_ptr).set_rect(rect);

            1
        }
        WM_DESTROY => {
            ExitProcess(0);
            1
        }
        _ => DefWindowProcW(hwnd, message, wparam, lparam),
    }
}

fn to_wstring(s: &str) -> Vec<u16> {
    OsStr::new(s)
        .encode_wide()
        .chain(Some(0).into_iter())
        .collect()
}

// unsafe fn from_wstring(wide: *const u16) -> OsString {
//     assert!(!wide.is_null());
//     for i in 0.. {
//         if *wide.offset(i) == 0 {
//             return OsStringExt::from_wide(std::slice::from_raw_parts(wide, i as usize));
//         }
//     }
//     unreachable!()
// }
