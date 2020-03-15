use winapi::shared::minwindef::*;
use winapi::shared::ntdef::*;
use winapi::shared::windef::*;
use winapi::shared::winerror::{self, FAILED, S_OK};
use winapi::um::errhandlingapi::*;
use winapi::um::libloaderapi::*;
use winapi::um::objidl::FORMATETC;
use winapi::um::ole2::*;
use winapi::um::winuser::*;
// use winapi::um::objidlbase::IEnumUnknown;
use com::{co_class, interfaces::iunknown::IUnknown, ComPtr};
use libc::c_void;
use std::cell::RefCell;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
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
    fn OleLockRunning(p_unknown: *mut c_void, f_lock: BOOL, f_last_unlock_closes: BOOL) -> HRESULT;

    fn ExitProcess(exit_code: UINT);
}

extern "system" {
    fn OleUninitialize();
}

#[co_class(implements(IOleClientSite, IOleInPlaceSite, IStorage))]
struct WebBrowser {
    hwnd_parent: HWND,
    rect_obj: RECT,
    ole_object: Option<ComPtr<dyn IOleObject>>,
    ole_in_place_object: RefCell<Option<ComPtr<dyn IOleInPlaceObject>>>,
    web_browser: Option<ComPtr<dyn IWebBrowser>>,
}

struct Userdata {
    h_instance: HINSTANCE,
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
        unsafe {
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

            let userdata = Box::new(Userdata { h_instance });
            let userdata = Box::into_raw(userdata);

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
                userdata as _, // TODO we need to call Box::into_raw when quiting so we do not leak
            );

            let mut web_browser = WebBrowser::allocate(
                handle,
                RECT {
                    left: 0,
                    top: 0,
                    right: 300,
                    bottom: 300,
                },
                None,
                RefCell::new(None),
                None,
            );

            let iole_client_site = web_browser
                .get_interface::<dyn IOleClientSite>()
                .expect("iole_client_site query failed");

            let istorage = web_browser
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

            // let hresult = ioleobject.set_client_site(iole_client_site.as_raw() as *mut c_void);

            // if FAILED(hresult) {
            //     panic!("set_client_site() failed");
            // }

            let hresult = OleSetContainedObject(ioleobject.as_raw() as _, 1);

            if FAILED(hresult) {
                panic!("OleSetContainedObject() failed");
            }

            web_browser.ole_object = Some(ioleobject.clone());

            let hresult = ioleobject.do_verb(
                -5,
                ptr::null_mut(),
                iole_client_site.as_raw() as _,
                -1,
                handle,
                &web_browser.rect_obj,
            );

            if FAILED(hresult) {
                panic!("ioleobject.do_verb() failed");
            }

            let iweb_browser = ioleobject
                .get_interface::<dyn IWebBrowser>()
                .expect("get interface IWebBrowser2 failed");

            web_browser.web_browser = Some(iweb_browser);

            // println!("yeet");
            // let hresult = ioleobject.set_client_site(iole_client_site);

            // if FAILED(hresult) {
            //     panic!("ioleobject.set_client_site() failed");
            // }

            //let iweb_browser = ioleobject.get_interface::<dyn IWebBrowser2>();
            // let mut iweb_browser = ptr::null_mut();
            // ioleobject.query_interface(&<dyn IWebBrowser2 as com::ComInterface>::IID, &mut iweb_browser);

            // eprintln!("{}", iweb_browser.is_null());

            ShowWindow(handle, SW_SHOWDEFAULT);

            web_browser
        }
    }

    fn set_rect(&self) {
        let mut rect: RECT = Default::default();
        unsafe {
            GetClientRect(self.hwnd_parent, &mut rect);
        }
        rect.top = 45;
        unsafe {
            self.ole_in_place_object
                .borrow()
                .as_ref()
                .unwrap()
                .set_object_rects(&rect, &rect);
        }
    }

    fn navigate(&self, url: &str) {
        let mut wstring = to_wstring(url);
        unsafe {
            self.web_browser.as_ref().unwrap().navigate(
                wstring.as_mut_ptr(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
            );
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

        let wb = WebBrowser::new();
        wb.navigate("http://google.com");

        let mut message: MSG = Default::default();
        while GetMessageW(&mut message, ptr::null_mut(), 0, 0) > 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
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
    match message {
        WM_CREATE => {
            let userdata: *mut Userdata = lparam as _;
            if userdata.is_null() {
                eprintln!("userdata is null");
                return DefWindowProcW(hwnd, message, wparam, lparam);
            }
            let userdata = userdata.as_mut().unwrap();

            println!("wm create");
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
                userdata.h_instance,
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
                userdata.h_instance,
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
                userdata.h_instance,
                ptr::null_mut(),
            );

            // let edit_handle = CreateWindowExW(
            //     0,
            //     to_wstring("EDIT").as_ptr(),
            //     to_wstring("http://google.com/").as_ptr(),
            //     WS_CHILD | WS_VISIBLE | WS_BORDER,
            //     260,
            //     10,
            //     200,
            //     20,
            //     hwnd,
            //     ptr::null_mut(),
            //     userdata.h_instance,
            //     lparam as _,
            // );

            // userdata.hwnd_addressbar = edit_handle;

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
                userdata.h_instance,
                ptr::null_mut(),
            );
            1
        }
        WM_COMMAND => {
            let cmd = LOWORD(wparam as _);
            match cmd {
                BTN_BACK => println!("go back"),
                BTN_NEXT => println!("go forward"),
                BTN_REFRESH => println!("refresh"),
                BTN_GO => {
                    println!("go");
                    // let mut buf: [u16; 1024] = [0; 1024];
                    // let userdata: *mut Userdata = lparam as _;
                    // let len =
                    //     GetWindowTextW((*userdata).hwnd_addressbar, buf.as_mut_ptr(), buf.len() as _)
                    //         as usize;
                    // if len == 0 {
                    //     return 1;
                    // }

                    // let s = OsString::from_wide(&buf[..len + 1]);
                    // println!("{:?}", s);
                }
                _ => {}
            }

            1
        }
        WM_SIZE => {
            println!("size");
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
