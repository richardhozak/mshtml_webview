#![allow(dead_code)]
#![allow(unused_variables)]

mod interface;
mod web_view;
mod window;

use std::{ffi::OsStr, os::windows::ffi::OsStrExt, ptr};

use winapi::{shared::windef::RECT, um::winuser::*};

use web_view::*;
use window::*;

fn main() {
    unsafe {
        let window = Window::new();

        let mut wb = WebView::new();
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
