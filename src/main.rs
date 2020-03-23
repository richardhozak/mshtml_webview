#![allow(dead_code)]
#![allow(unused_variables)]

use com::{co_class, interfaces::iunknown::IUnknown, ComPtr, ComRc};
use libc::c_void;
use std::ffi::{OsStr, OsString};
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::ptr;
use winapi::shared::guiddef::IID_NULL;
use winapi::shared::minwindef::{BOOL, DWORD, LOWORD, LPARAM, LRESULT, UINT, WORD, WPARAM};
use winapi::shared::ntdef::{LOCALE_SYSTEM_DEFAULT, LONG};
use winapi::shared::windef::{HWND, RECT};
use winapi::shared::winerror::{self, FAILED, HRESULT, S_OK};
use winapi::shared::wtypes::{VT_BSTR, VT_VARIANT};
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::libloaderapi::GetModuleHandleW;
use winapi::um::oaidl::{DISPID, DISPPARAMS, VARIANT};
use winapi::um::objidl::FORMATETC;
use winapi::um::ole2::OleInitialize;
use winapi::um::oleauto::{
    SafeArrayAccessData, SafeArrayCreateVector, SafeArrayDestroy, SysAllocString, SysFreeString,
};
use winapi::um::winuser::*;

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

#[co_class(implements(IOleClientSite, IOleInPlaceSite, IStorage, IDocHostUIHandler))]
struct WebBrowser {
    inner: Option<WebBrowserInner>,
}

struct WebBrowserInner {
    hwnd_parent: HWND,
    rect: RECT,
    ole_in_place_object: ComPtr<dyn IOleInPlaceObject>,
    web_browser: ComPtr<dyn IWebBrowser>,
    invoke_receiver: *mut ExternalInvokeReceiver, // this should be ComPtr
}

#[co_class(implements(IDispatch))]
struct ExternalInvokeReceiver {
    h_wnd: HWND,
}

const INVOKE_CALLBACK_MSG: UINT = WM_USER + 1;

impl ExternalInvokeReceiver {
    fn new() -> Box<ExternalInvokeReceiver> {
        ExternalInvokeReceiver::allocate(ptr::null_mut())
    }

    fn set_target(&mut self, h_wnd: HWND) {
        self.h_wnd = h_wnd;
    }

    fn invoke_callback(&self, data: String) {
        println!("invoke callback");
        let data = Box::new(data);
        let data = Box::into_raw(data);
        unsafe {
            SendMessageW(
                self.h_wnd,
                INVOKE_CALLBACK_MSG,
                0,
                std::mem::transmute(data),
            );
        }
    }
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

    fn write(&self, document: &str) {
        println!("writing {}", document);
        let inner = self.inner.as_ref().unwrap();
        unsafe {
            let mut document_dispatch = ptr::null_mut::<c_void>();
            let h_result = inner.web_browser.get_document(&mut document_dispatch);
            if FAILED(h_result) || document_dispatch.is_null() {
                panic!("get_document failed {}", h_result);
            }

            let document_dispatch =
                ComRc::<dyn IDispatch>::from_raw(document_dispatch as *mut *mut _);

            let html_document2 = document_dispatch
                .get_interface::<dyn IHTMLDocument2>()
                .expect("cannot get IHTMLDocument2 interface");

            let safe_array = SafeArrayCreateVector(VT_VARIANT as _, 0, 1);
            if safe_array.is_null() {
                panic!("SafeArrayCreate failed");
            }
            let mut data: [*mut VARIANT; 1] = [ptr::null_mut()];
            let h_result = SafeArrayAccessData(safe_array, data.as_mut_ptr() as _);
            if FAILED(h_result) {
                panic!("SafeArrayAccessData failed");
            }

            let document = to_wstring(document);
            let document = SysAllocString(document.as_ptr());

            if document.is_null() {
                panic!("SysAllocString document failed");
            }
            let variant = &mut (*data[0]);
            variant.n1.n2_mut().vt = VT_BSTR as _;
            *variant.n1.n2_mut().n3.bstrVal_mut() = document;
            if FAILED(html_document2.write(safe_array)) {
                panic!("html_document2.write() failed");
            }
            if FAILED(html_document2.close()) {
                panic!("html_document2.close() failed");
            }

            SysFreeString(document);
            SafeArrayDestroy(safe_array);
        }
    }

    fn eval(&self, js: &str) {
        let inner = self.inner.as_ref().unwrap();
        unsafe {
            let mut document_dispatch = ptr::null_mut::<c_void>();
            let h_result = inner.web_browser.get_document(&mut document_dispatch);
            if FAILED(h_result) || document_dispatch.is_null() {
                panic!("get_document failed {}", h_result);
            }

            let document_dispatch =
                ComRc::<dyn IDispatch>::from_raw(document_dispatch as *mut *mut _);

            let html_document = document_dispatch
                .get_interface::<dyn IHTMLDocument>()
                .expect("cannot get IHTMLDocument interface");

            let mut script_dispatch = ptr::null_mut::<c_void>();
            let h_result = html_document.get_script(&mut script_dispatch);
            if FAILED(h_result) || script_dispatch.is_null() {
                panic!("get_script failed {}", h_result);
            }

            let script_dispatch = ComRc::<dyn IDispatch>::from_raw(script_dispatch as *mut *mut _);

            let mut eval_name = to_wstring("eval");
            let mut names = [eval_name.as_mut_ptr()];
            let mut disp_ids = [DISPID::default()];
            assert_eq!(names.len(), disp_ids.len());

            // get_ids_of_names can fail if there is no loaded document
            // should we hande this by loading empty document?

            let h_result = script_dispatch.get_ids_of_names(
                &IID_NULL,
                names.as_mut_ptr(),
                names.len() as _,
                LOCALE_SYSTEM_DEFAULT,
                disp_ids.as_mut_ptr(),
            );
            if FAILED(h_result) {
                panic!("get_ids_of_names failed {}", h_result);
            }

            let js = to_wstring(js);

            // we need to free this later
            // with SysFreeString
            let js = SysAllocString(js.as_ptr());
            let mut varg = VARIANT::default();
            varg.n1.n2_mut().vt = VT_BSTR as _;

            // we cannot pass regular BSTR here,
            // it needs to be allocated with SysAllocString
            // which also allocates inner data for
            // the specific BSTR such as refcount
            *varg.n1.n2_mut().n3.bstrVal_mut() = js;

            let mut args = [varg];
            let mut disp_params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),          // array of positional arguments
                rgdispidNamedArgs: ptr::null_mut(), // array of dispids for named args
                cArgs: args.len() as _,             // number of position arguments
                cNamedArgs: 0,                      // number of named args - none
            };
            let h_result = script_dispatch.invoke(
                disp_ids[0],
                &IID_NULL,
                0,
                1,
                &mut disp_params,
                ptr::null_mut(), // should we implement result?
                ptr::null_mut(),
                ptr::null_mut(),
            );

            SysFreeString(js);

            // this should be catchable by user,
            // it does not always have to be irrecoverable error
            if FAILED(h_result) {
                panic!("invoke failed {}", h_result);
            }
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

            let mut invoke_receiver = ExternalInvokeReceiver::new();
            invoke_receiver.set_target(h_wnd);
            let invoke_receiver = Box::into_raw(invoke_receiver);

            self.inner = Some(WebBrowserInner {
                hwnd_parent: h_wnd,
                rect,
                ole_in_place_object,
                web_browser,
                invoke_receiver,
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

struct Window {
    h_wnd: HWND,
    fullscreen: bool,
    saved_style: LONG,
    saved_ex_style: LONG,
    saved_rect: RECT,
}

impl Window {
    fn new() -> Self {
        // TODO: move some of this logic into some sort of event loop or main application
        // the idea is to have application, that can spawn windows and webviews that
        // can be spawned multiple times into these windows

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

            Window {
                h_wnd,
                fullscreen: false,
                saved_style: 0,
                saved_ex_style: 0,
                saved_rect: Default::default(),
            }
        }
    }

    fn handle(&self) -> HWND {
        self.h_wnd
    }

    fn save_style(&mut self) {
        unsafe {
            self.saved_style = GetWindowLongW(self.h_wnd, GWL_STYLE);
            self.saved_ex_style = GetWindowLongW(self.h_wnd, GWL_EXSTYLE);
            GetWindowRect(self.h_wnd, &mut self.saved_rect);
        }
    }

    fn restore_style(&self) {
        unsafe {
            SetWindowLongW(self.h_wnd, GWL_STYLE, self.saved_style);
            SetWindowLongW(self.h_wnd, GWL_EXSTYLE, self.saved_ex_style);
            let rect = &self.saved_rect;
            SetWindowPos(
                self.h_wnd,
                ptr::null_mut(),
                rect.left,
                rect.top,
                rect.right - rect.left,
                rect.bottom - rect.top,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            );
        }
    }

    fn set_fullscreen(&mut self, fullscreen: bool) {
        if self.fullscreen == fullscreen {
            return;
        }

        if !self.fullscreen {
            self.save_style();
        }

        self.fullscreen = fullscreen;

        if !self.fullscreen {
            self.restore_style();
            return;
        }

        unsafe {
            let mut monitor_info: MONITORINFO = Default::default();
            monitor_info.cbSize = std::mem::size_of::<MONITORINFO>() as _;
            GetMonitorInfoW(
                MonitorFromWindow(self.h_wnd, MONITOR_DEFAULTTONEAREST),
                &mut monitor_info,
            );

            SetWindowLongW(
                self.h_wnd,
                GWL_STYLE,
                self.saved_style & !(WS_CAPTION | WS_THICKFRAME) as LONG,
            );

            SetWindowLongW(
                self.h_wnd,
                GWL_EXSTYLE,
                self.saved_ex_style
                    & !(WS_EX_DLGMODALFRAME
                        | WS_EX_WINDOWEDGE
                        | WS_EX_CLIENTEDGE
                        | WS_EX_STATICEDGE) as LONG,
            );

            let rect = &monitor_info.rcMonitor;
            SetWindowPos(
                self.h_wnd,
                ptr::null_mut(),
                rect.left,
                rect.top,
                rect.right - rect.left,
                rect.bottom - rect.top,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            );
        }
    }
}

fn main() {
    unsafe {
        let window = Window::new();

        let mut wb = WebBrowser::new();
        wb.initialize(
            window.handle(),
            RECT {
                left: 0,
                right: 300,
                top: 45,
                bottom: 300,
            },
        );
        wb.navigate("about:blank");

        let wb_ptr = Box::into_raw(wb);

        SetWindowLongPtrW(window.handle(), GWLP_USERDATA, std::mem::transmute(wb_ptr));
        ShowWindow(window.handle(), SW_SHOWDEFAULT);

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
const BTN_EVAL: WORD = 5;
const BTN_WRITE_DOC: WORD = 6;

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
                to_wstring("http://example.com/").as_ptr(),
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

            CreateWindowExW(
                0,
                to_wstring("BUTTON").as_ptr(),
                to_wstring("Eval").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                665,
                5,
                80,
                30,
                hwnd,
                BTN_EVAL as _,
                h_instance,
                ptr::null_mut(),
            );

            CreateWindowExW(
                0,
                to_wstring("BUTTON").as_ptr(),
                to_wstring("Write").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                750,
                5,
                80,
                30,
                hwnd,
                BTN_WRITE_DOC as _,
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
                BTN_EVAL => {
                    (*wb_ptr).eval("external.invoke('test');");
                    // (*wb_ptr).eval("alert('hello');");
                }
                BTN_WRITE_DOC => {
                    (*wb_ptr).write("<p>Hello world!</p>");
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
        INVOKE_CALLBACK_MSG => {
            let wb_ptr: *mut WebBrowser =
                std::mem::transmute(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
            if wb_ptr.is_null() {
                return 1;
            }

            let data: *mut String = std::mem::transmute(lparam);
            let data = Box::from_raw(data);
            println!("got data {}", data);

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
