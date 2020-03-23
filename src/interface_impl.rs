use super::interface::*;
use super::{ExternalInvokeReceiver, WebBrowser};

use com::interfaces::IUnknown;
use libc::c_void;
use std::ffi::OsString;
use winapi::shared::guiddef::IID;
use winapi::shared::minwindef::UINT;
use winapi::shared::minwindef::WORD;
use winapi::shared::ntdef::LCID;
use winapi::shared::winerror::HRESULT;
use winapi::shared::winerror::{E_FAIL, E_NOINTERFACE, E_NOTIMPL, E_PENDING, S_FALSE, S_OK};
use winapi::shared::wtypes::VARTYPE;
use winapi::shared::wtypes::VT_BSTR;
use winapi::shared::wtypesbase::LPOLESTR;
use winapi::um::oaidl::DISPID;
use winapi::um::oaidl::DISPPARAMS;
use winapi::um::oaidl::EXCEPINFO;
use winapi::um::oaidl::VARIANT;

use std::ptr;
use std::ptr::NonNull;

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
        if self.inner.is_none() {
            *phwnd = ptr::null_mut();
            return E_PENDING;
        }

        *phwnd = self.inner.as_ref().unwrap().hwnd_parent;
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
        S_OK
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
        *pp_frame = ptr::null_mut();
        *pp_doc = ptr::null_mut();
        *lprc_pos_rect = self.inner.as_ref().unwrap().rect;
        *lprc_clip_rect = *lprc_pos_rect;

        (*lp_frame_info).fMDIApp = 0;
        (*lp_frame_info).hwndFrame = self.inner.as_ref().unwrap().hwnd_parent;
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

impl IDocHostUIHandler for WebBrowser {
    unsafe fn show_context_menu(
        &self,
        _: u32,
        _: *mut winapi::shared::windef::POINT,
        _: *mut core::ffi::c_void,
        _: *mut core::ffi::c_void,
    ) -> i32 {
        S_OK
    }
    unsafe fn get_host_info(&self, _: *mut core::ffi::c_void) -> i32 {
        E_NOTIMPL
    }
    unsafe fn show_ui(
        &self,
        _: u32,
        _: *mut core::ffi::c_void,
        _: *mut core::ffi::c_void,
        _: *mut core::ffi::c_void,
        _: *mut core::ffi::c_void,
    ) -> i32 {
        S_OK
    }
    unsafe fn hide_ui(&self) -> i32 {
        S_OK
    }
    unsafe fn update_ui(&self) -> i32 {
        S_OK
    }
    unsafe fn enable_modeless(&self, _: i32) -> i32 {
        S_OK
    }
    unsafe fn on_doc_window_activate(&self, _: i32) -> i32 {
        S_OK
    }
    unsafe fn on_frame_window_activate(&self, _: i32) -> i32 {
        S_OK
    }
    unsafe fn resize_border(
        &self,
        _: *const winapi::shared::windef::RECT,
        _: *mut core::ffi::c_void,
        _: i32,
    ) -> i32 {
        S_OK
    }
    unsafe fn translate_accelerator(
        &self,
        _: *mut winapi::um::winuser::MSG,
        _: *const winapi::shared::guiddef::GUID,
        _: u32,
    ) -> i32 {
        S_FALSE
    }
    unsafe fn get_option_key_path(&self, _: *mut *mut u16, _: u32) -> i32 {
        S_FALSE
    }
    unsafe fn get_drop_target(
        &self,
        _: *mut core::ffi::c_void,
        _: *mut *mut core::ffi::c_void,
    ) -> i32 {
        S_FALSE
    }
    unsafe fn get_external(&self, external: *mut *mut core::ffi::c_void) -> i32 {
        println!("get external");
        let inner = self.inner.as_ref().unwrap();
        (*inner.invoke_receiver).add_ref();
        *external = inner.invoke_receiver as _;
        S_OK
    }
    unsafe fn translate_url(&self, _: u32, _: *mut u16, ppch_url_out: *mut *mut u16) -> i32 {
        *ppch_url_out = ptr::null_mut();
        S_FALSE
    }
    unsafe fn filter_data_object(
        &self,
        _: *mut core::ffi::c_void,
        pp_do_ret: *mut *mut core::ffi::c_void,
    ) -> i32 {
        *pp_do_ret = ptr::null_mut();
        S_FALSE
    }
}

unsafe fn from_wstring(ptr: *const u16) -> OsString {
    use std::os::windows::ffi::OsStringExt;

    let len = (0..).take_while(|&i| *ptr.offset(i) != 0).count();
    let slice = std::slice::from_raw_parts(ptr, len);

    OsString::from_wide(slice)
}

unsafe fn from_utf16(ptr: *const u16) -> String {
    let len = (0..).take_while(|&i| *ptr.offset(i) != 0).count();
    let slice = std::slice::from_raw_parts(ptr, len);
    String::from_utf16(slice).expect("invalid utf16")
}

const WEBVIEW_JS_INVOKE_ID: DISPID = 0x1000;

impl IDispatch for ExternalInvokeReceiver {
    unsafe fn get_type_info_count(&self, pctinfo: *mut UINT) -> HRESULT {
        S_OK
    }
    unsafe fn get_type_info(
        &self,
        i_ti_info: UINT,
        icid: LCID,
        pp_ti_info: *mut *mut c_void,
    ) -> HRESULT {
        S_OK
    }
    unsafe fn get_ids_of_names(
        &self,
        riid: *const IID,
        rgsz_names: *mut LPOLESTR,
        c_names: UINT,
        lcid: LCID,
        rg_disp_id: *mut DISPID,
    ) -> HRESULT {
        let names = std::slice::from_raw_parts(rgsz_names, c_names as _);
        if names.len() == 1 {
            let name = from_wstring(names[0]);
            if name == "invoke" {
                // map the invoke function on external object to this id
                *rg_disp_id.offset(0) = WEBVIEW_JS_INVOKE_ID;
                return S_OK;
            }
        }

        S_FALSE
    }
    unsafe fn invoke(
        &self,
        disp_id_member: DISPID,
        riid: *const IID,
        lcid: LCID,
        w_flags: WORD,
        p_disp_params: *mut DISPPARAMS,
        p_var_result: *mut VARIANT,
        p_excep_info: *mut EXCEPINFO,
        pu_arg_err: *mut UINT,
    ) -> HRESULT {
        // first we check if the message the webview is trying to
        // invoke is the method we gave it in get_ids_of_names
        // through the custom id we specified
        if disp_id_member == WEBVIEW_JS_INVOKE_ID {
            let params = NonNull::new(p_disp_params).expect("p_disp_params is null");
            let params = params.as_ref();
            let vargs = std::slice::from_raw_parts(params.rgvarg, params.cArgs as _);

            // we only handle invoke function which has only one positional argument
            // and the argument needs to be string
            if vargs.len() == 1 {
                let varg = &vargs[0];

                // check if the argument is string,
                // convert it to String from utf16
                // and pass it further
                if varg.n1.n2().vt == VT_BSTR as VARTYPE {
                    let arg = *varg.n1.n2().n3.bstrVal();
                    let arg = from_utf16(arg);
                    self.invoke_callback(arg);
                    return S_OK;
                }
            }
        }

        S_FALSE
    }
}
