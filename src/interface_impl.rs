use super::WebBrowser;
use super::interface::*;
use super::OleLockRunning;

use winapi::shared::winerror::{E_FAIL, E_NOINTERFACE, E_NOTIMPL, FAILED, S_OK};
use com::{ComPtr, interfaces::IUnknown};

use std::ptr;

impl IOleClientSite for WebBrowser {
    unsafe fn save_object(&self) -> i32 {
        E_NOTIMPL
    }
    unsafe fn get_moniker(
        &self,
        dw_assign: u32,
        dw_which_moniker: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        // dw_assign: OLEGETMONIKER_ONLYIFTHERE = 1
        // dw_which_moniker: OLEWHICHMK_CONTAINER = 1

        if dw_assign == 1 || dw_which_moniker == 1 {
            eprintln!("assign faield");
            E_FAIL
        } else {
            E_NOTIMPL
        }
    }
    unsafe fn get_container(&self, _: *mut *mut std::ffi::c_void) -> i32 {
        E_NOINTERFACE
    }
    unsafe fn show_object(&self) -> i32 {
        S_OK
    }
    unsafe fn on_show_window(&self, _: i32) -> i32 {
        S_OK
    }
    unsafe fn request_new_object_layout(&self) -> i32 {
        E_NOTIMPL
    }
}

impl IOleWindow for WebBrowser {
    unsafe fn get_window(&self, phwnd: *mut *mut winapi::shared::windef::HWND__) -> i32 {
        *phwnd = self.hwnd_parent;
        S_OK
    }
    unsafe fn context_sensitive_help(&self, _: i32) -> i32 {
        E_NOTIMPL
    }
}

impl IOleInPlaceSite for WebBrowser {
    unsafe fn can_in_place_activate(&self) -> i32 {
        S_OK
    }
    unsafe fn on_in_place_activate(&self) -> i32 {
        eprintln!("on_in_place_activate");

        let ole_object = self
            .ole_object
            .as_ref()
            .expect("webbrowser incorrectly initialized, ole_object is not present");

        OleLockRunning(ole_object.as_raw() as _, 1, 0);
        let mut ole_in_place_object = ptr::null_mut();
        let result = ole_object.query_interface(
            &<dyn IOleInPlaceObject as com::ComInterface>::IID,
            &mut ole_in_place_object,
        );

        if FAILED(result) {
            panic!("cannot query ole_in_place_object");
        }

        let ole_in_place_object =
            ComPtr::<dyn IOleInPlaceObject>::new(std::mem::transmute(ole_in_place_object));

        ole_in_place_object.set_object_rects(&self.rect_obj, &self.rect_obj);

        self.ole_in_place_object.set(Some(ole_in_place_object));

        // implement oleinplaceobject query interface
        unimplemented!()
    }
    unsafe fn on_ui_activate(&self) -> i32 {
        S_OK
    }
    unsafe fn get_window_context(
        &self,
        pp_frame: *mut *mut std::ffi::c_void,
        pp_doc: *mut *mut std::ffi::c_void,
        lprc_pos_rect: *mut winapi::shared::windef::RECT,
        lprc_clip_rect: *mut winapi::shared::windef::RECT,
        lp_frame_info: *mut OLEINPLACEFRAMEINFO,
    ) -> i32 {
        eprintln!("get window context");
        *pp_frame = ptr::null_mut();
        *pp_doc = ptr::null_mut();
        (*lprc_pos_rect).left = self.rect_obj.left;
        (*lprc_pos_rect).top = self.rect_obj.top;
        (*lprc_pos_rect).right = self.rect_obj.right;
        (*lprc_pos_rect).bottom = self.rect_obj.bottom;
        *lprc_clip_rect = *lprc_pos_rect;

        (*lp_frame_info).fMDIApp = 0;
        (*lp_frame_info).hwndFrame = self.hwnd_parent;
        (*lp_frame_info).haccel = ptr::null_mut();
        (*lp_frame_info).cAccelEntries = 0;
        S_OK
    }
    unsafe fn scroll(&self, _: winapi::shared::windef::SIZE) -> i32 {
        E_NOTIMPL
    }
    unsafe fn on_ui_deactivate(&self, _: i32) -> i32 {
        S_OK
    }
    unsafe fn on_in_place_deactivate(&self) -> i32 {
        // implement null fields
        S_OK
    }
    unsafe fn discard_undo_state(&self) -> i32 {
        E_NOTIMPL
    }
    unsafe fn deactivate_and_undo(&self) -> i32 {
        E_NOTIMPL
    }
    unsafe fn on_pos_rect_change(&self, _: *mut winapi::shared::windef::RECT) -> i32 {
        E_NOTIMPL
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
        E_NOTIMPL
    }
    unsafe fn open_stream(
        &self,
        _: *const u16,
        _: *mut std::ffi::c_void,
        _: u32,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        E_NOTIMPL
    }
    unsafe fn create_storage(
        &self,
        _: *const u16,
        _: u32,
        _: u32,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        E_NOTIMPL
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
        E_NOTIMPL
    }
    unsafe fn copy_to(
        &self,
        _: u32,
        _: *const winapi::shared::guiddef::GUID,
        _: *const *const u16,
        _: *mut std::ffi::c_void,
    ) -> i32 {
        E_NOTIMPL
    }
    unsafe fn move_element_to(
        &self,
        _: *const u16,
        _: *mut std::ffi::c_void,
        _: *const u16,
        _: u32,
    ) -> i32 {
        E_NOTIMPL
    }
    unsafe fn commit(&self, _: u32) -> i32 {
        E_NOTIMPL
    }
    unsafe fn revert(&self) -> i32 {
        E_NOTIMPL
    }
    unsafe fn enum_elements(
        &self,
        _: u32,
        _: *mut std::ffi::c_void,
        _: u32,
        _: *mut *mut std::ffi::c_void,
    ) -> i32 {
        E_NOTIMPL
    }
    unsafe fn destroy_element(&self, _: *const u16) -> i32 {
        E_NOTIMPL
    }
    unsafe fn rename_element(&self, _: *const u16, _: *const u16) -> i32 {
        E_NOTIMPL
    }
    unsafe fn set_element_times(
        &self,
        _: *const u16,
        _: *const winapi::shared::minwindef::FILETIME,
        _: *const winapi::shared::minwindef::FILETIME,
        _: *const winapi::shared::minwindef::FILETIME,
    ) -> i32 {
        E_NOTIMPL
    }
    unsafe fn set_class(&self, _: *const winapi::shared::guiddef::GUID) -> i32 {
        S_OK
    }
    unsafe fn set_state_bits(&self, _: u32, _: u32) -> i32 {
        E_NOTIMPL
    }
    unsafe fn stat(&self, _: *mut winapi::um::objidlbase::STATSTG, _: u32) -> i32 {
        E_NOTIMPL
    }
}
