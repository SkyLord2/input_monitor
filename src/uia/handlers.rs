use std::ffi::c_void;
use std::mem::MaybeUninit;
use std::sync::atomic::{AtomicU32, Ordering};

use windows::core::{BSTR, GUID, HRESULT, IUnknown, IUnknown_Vtbl, VARIANT, Interface};
use windows::Win32::Foundation::{E_NOINTERFACE, S_OK};
use windows::Win32::UI::Accessibility::{
    IUIAutomationElement, IUIAutomationEventHandler, IUIAutomationEventHandler_Vtbl,
    IUIAutomationFocusChangedEventHandler, IUIAutomationFocusChangedEventHandler_Vtbl,
    IUIAutomationPropertyChangedEventHandler, IUIAutomationPropertyChangedEventHandler_Vtbl,
    UIA_DocumentControlTypeId, UIA_EditControlTypeId, UIA_GroupControlTypeId,
    UIA_Text_TextChangedEventId, UIA_ValueValuePropertyId, UIA_EVENT_ID, UIA_PROPERTY_ID,
};

use crate::debounce::debounce_print;
use crate::uia::text::get_text_deep;

#[repr(C)]
pub struct ManualFocusHandler {
    vtable: *const IUIAutomationFocusChangedEventHandler_Vtbl,
    ref_count: AtomicU32,
}

impl ManualFocusHandler {
    const VTABLE: IUIAutomationFocusChangedEventHandler_Vtbl =
        IUIAutomationFocusChangedEventHandler_Vtbl {
            base__: IUnknown_Vtbl {
                QueryInterface: Self::query_interface,
                AddRef: Self::add_ref,
                Release: Self::release,
            },
            HandleFocusChangedEvent: Self::handle_focus_changed_event,
        };

    pub fn new() -> *mut ManualFocusHandler {
        let handler = Box::new(Self {
            vtable: &Self::VTABLE,
            ref_count: AtomicU32::new(1),
        });
        Box::into_raw(handler)
    }

    unsafe extern "system" fn query_interface(
        this: *mut c_void,
        iid: *const GUID,
        interface: *mut *mut c_void,
    ) -> HRESULT {
        let this = this as *mut Self;
        unsafe {
            if *iid == IUnknown::IID || *iid == IUIAutomationFocusChangedEventHandler::IID {
                *interface = this as *mut c_void;
                Self::add_ref(this as *mut c_void);
                S_OK
            } else {
                *interface = std::ptr::null_mut();
                E_NOINTERFACE
            }
        }
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        unsafe { (*(this as *mut Self)).ref_count.fetch_add(1, Ordering::Relaxed) + 1 }
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        unsafe {
            let this = this as *mut Self;
            let count = (*this).ref_count.fetch_sub(1, Ordering::Relaxed) - 1;
            if count == 0 {
                let _ = Box::from_raw(this);
            }
            count
        }
    }

    unsafe extern "system" fn handle_focus_changed_event(
        _this: *mut c_void,
        sender: *mut c_void,
    ) -> HRESULT {
        unsafe {
            if !sender.is_null() {
                let element: &IUIAutomationElement = std::mem::transmute(&sender);
                if let Ok(control_type) = element.CurrentControlType() {
                    if control_type == UIA_EditControlTypeId || control_type == UIA_GroupControlTypeId {
                        let name = element.CurrentName().unwrap_or(BSTR::new());
                        println!(">>> [焦点切换] 控件类型: {:?}, 进入输入框: '{}'", control_type, name);
                        if let Ok(text) = get_text_deep(element) {
                            println!("    当前内容: {}", text);
                        }
                    }
                }
            }
        }
        S_OK
    }
}

#[repr(C)]
pub struct ManualPropertyHandler {
    vtable: *const IUIAutomationPropertyChangedEventHandler_Vtbl,
    ref_count: AtomicU32,
}

impl ManualPropertyHandler {
    const VTABLE: IUIAutomationPropertyChangedEventHandler_Vtbl =
        IUIAutomationPropertyChangedEventHandler_Vtbl {
            base__: IUnknown_Vtbl {
                QueryInterface: Self::query_interface,
                AddRef: Self::add_ref,
                Release: Self::release,
            },
            HandlePropertyChangedEvent: Self::handle_property_changed_event,
        };

    pub fn new() -> *mut ManualPropertyHandler {
        let handler = Box::new(Self {
            vtable: &Self::VTABLE,
            ref_count: AtomicU32::new(1),
        });
        Box::into_raw(handler)
    }

    unsafe extern "system" fn query_interface(
        this: *mut c_void,
        iid: *const GUID,
        interface: *mut *mut c_void,
    ) -> HRESULT {
        let this = this as *mut Self;
        unsafe {
            if *iid == IUnknown::IID || *iid == IUIAutomationPropertyChangedEventHandler::IID {
                *interface = this as *mut c_void;
                Self::add_ref(this as *mut c_void);
                S_OK
            } else {
                *interface = std::ptr::null_mut();
                E_NOINTERFACE
            }
        }
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        unsafe { (*(this as *mut Self)).ref_count.fetch_add(1, Ordering::Relaxed) + 1 }
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        unsafe {
            let this = this as *mut Self;
            let count = (*this).ref_count.fetch_sub(1, Ordering::Relaxed) - 1;
            if count == 0 {
                let _ = Box::from_raw(this);
            }
            count
        }
    }

    unsafe extern "system" fn handle_property_changed_event(
        _this: *mut c_void,
        sender: *mut c_void,
        property_id: UIA_PROPERTY_ID,
        new_value: MaybeUninit<VARIANT>,
    ) -> HRESULT {
        if property_id == UIA_ValueValuePropertyId && !sender.is_null() {
            unsafe {
                let element: &IUIAutomationElement = std::mem::transmute(&sender);

                if let Ok(control_type) = element.CurrentControlType() {
                    if control_type == UIA_EditControlTypeId || control_type == UIA_GroupControlTypeId {
                        let value_ref = new_value.assume_init_ref();
                        let val_str = value_ref.to_string();
                        let current_text = if val_str.is_empty() {
                            get_text_deep(element).unwrap_or_default()
                        } else {
                            val_str
                        };

                        let name = element.CurrentName().unwrap_or(BSTR::new());
                        if element.CurrentHasKeyboardFocus().unwrap_or_default().as_bool() {
                            if !current_text.is_empty() {
                                debounce_print(format!(
                                    "    [输入监测] 控件类型: {:?}, '{}' 变更为: {}",
                                    control_type, name, current_text
                                ));
                            }
                        }
                    }
                }
            }
        }
        S_OK
    }
}

#[repr(C)]
pub struct ManualTextChangedHandler {
    vtable: *const IUIAutomationEventHandler_Vtbl,
    ref_count: AtomicU32,
}

impl ManualTextChangedHandler {
    const VTABLE: IUIAutomationEventHandler_Vtbl = IUIAutomationEventHandler_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: Self::query_interface,
            AddRef: Self::add_ref,
            Release: Self::release,
        },
        HandleAutomationEvent: Self::handle_automation_event,
    };

    pub fn new() -> *mut ManualTextChangedHandler {
        let handler = Box::new(Self {
            vtable: &Self::VTABLE,
            ref_count: AtomicU32::new(1),
        });
        Box::into_raw(handler)
    }

    unsafe extern "system" fn query_interface(
        this: *mut c_void,
        iid: *const GUID,
        interface: *mut *mut c_void,
    ) -> HRESULT {
        let this = this as *mut Self;
        unsafe {
            if *iid == IUnknown::IID || *iid == IUIAutomationEventHandler::IID {
                *interface = this as *mut c_void;
                Self::add_ref(this as *mut c_void);
                S_OK
            } else {
                *interface = std::ptr::null_mut();
                E_NOINTERFACE
            }
        }
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        unsafe { (*(this as *mut Self)).ref_count.fetch_add(1, Ordering::Relaxed) + 1 }
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        unsafe {
            let this = this as *mut Self;
            let count = (*this).ref_count.fetch_sub(1, Ordering::Relaxed) - 1;
            if count == 0 {
                let _ = Box::from_raw(this);
            }
            count
        }
    }

    unsafe extern "system" fn handle_automation_event(
        _this: *mut c_void,
        sender: *mut c_void,
        event_id: UIA_EVENT_ID,
    ) -> HRESULT {
        if event_id != UIA_Text_TextChangedEventId {
            return S_OK;
        }
        unsafe {
            if !sender.is_null() {
                let element: &IUIAutomationElement = std::mem::transmute(&sender);
                if !element.CurrentHasKeyboardFocus().unwrap_or_default().as_bool() {
                    return S_OK;
                }
                let control_type = element.CurrentControlType().unwrap_or_default();
                if control_type != UIA_EditControlTypeId
                    && control_type != UIA_DocumentControlTypeId
                    && control_type != UIA_GroupControlTypeId
                {
                    return S_OK;
                }
                if let Ok(text) = get_text_deep(element) {
                    if !text.is_empty() {
                        let name = element.CurrentName().unwrap_or(BSTR::new());
                        debounce_print(format!(
                            "    [输入监测] (TextChanged) {:?}, '{}' 变更为: {}",
                            control_type, name, text
                        ));
                    }
                }
            }
        }
        S_OK
    }
}
