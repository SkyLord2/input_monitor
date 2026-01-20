use std::cell::RefCell;
use std::ffi::c_void;
use std::mem::MaybeUninit;
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::mpsc::{self, Sender};
use std::sync::Arc;
use std::sync::OnceLock;
use std::thread;
use std::time::Duration;

use windows::core::{Interface, Result, BSTR, GUID, HRESULT, IUnknown, IUnknown_Vtbl, VARIANT};
use windows::Win32::Foundation::{E_NOINTERFACE, S_OK};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, 
    CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED,
};
use windows::Win32::System::Variant::VT_I4;
use windows::Win32::System::Ole::{
    SafeArrayCreateVector, SafeArrayDestroy, SafeArrayPutElement,
};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationElement, IUIAutomationFocusChangedEventHandler,
    IUIAutomationFocusChangedEventHandler_Vtbl, IUIAutomationPropertyChangedEventHandler,
    IUIAutomationPropertyChangedEventHandler_Vtbl, IUIAutomationValuePattern,
    IUIAutomationCondition,
    IUIAutomationEventHandler, IUIAutomationEventHandler_Vtbl,
    TreeScope_Descendants, UIA_EditControlTypeId, UIA_ValuePatternId, UIA_ValueValuePropertyId,
    IUIAutomationTextPattern, UIA_TextPatternId, UIA_GroupControlTypeId,
    UIA_DocumentControlTypeId,
    UIA_IsTextPatternAvailablePropertyId, UIA_IsValuePatternAvailablePropertyId,
    UIA_Text_TextChangedEventId,
};

thread_local! {
    static TL_AUTOMATION: RefCell<Option<IUIAutomation>> = const { RefCell::new(None) };
}

const MAX_TEXT_LEN: i32 = 4096;
const DEBOUNCE_MS: u64 = 200;

struct DebounceEvent {
    message: String,
}

static DEBOUNCE_SENDER: OnceLock<Sender<DebounceEvent>> = OnceLock::new();

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).ok()?;
    }

    println!("初始化 UIA (手动实现 COM 模式)...");
    let automation: IUIAutomation = unsafe {
        CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)?
    };

    // 1. 创建手动实现的 Handler
    let raw_focus_handler = ManualFocusHandler::new();
    let focus_interface: IUIAutomationFocusChangedEventHandler =
        unsafe { std::mem::transmute(raw_focus_handler) };

    let raw_prop_handler = ManualPropertyHandler::new();
    let prop_interface: IUIAutomationPropertyChangedEventHandler =
        unsafe { std::mem::transmute(raw_prop_handler) };

    let raw_text_handler = ManualTextChangedHandler::new();
    let text_interface: IUIAutomationEventHandler =
        unsafe { std::mem::transmute(raw_text_handler) };

    unsafe {
        // 注册焦点监听
        automation.AddFocusChangedEventHandler(None, &focus_interface)?;
        println!("  [+] 焦点监听已注册");

        let root = automation.GetRootElement()?;
        
        // [修复核心] 创建一个空的 SAFEARRAY，而不是传 NULL
        // 参数: 元素类型 VT_I4 (整数), 下界 0, 元素数量 0
        let sa = SafeArrayCreateVector(VT_I4, 0, 1);

        if sa.is_null() {
            panic!("无法创建 SAFEARRAY: 内存不足");
        }

        // 3. 填充 SAFEARRAY
        //    注意：SafeArrayPutElement 需要指针。
        //    UIA_ValueValuePropertyId 是一个 struct(u32)，我们需要转成 i32 放入
        let idx: i32 = 0; // 数组索引
        let prop_id_val = UIA_ValueValuePropertyId.0 as i32; // 属性值
        
        // void* 指针转换戏法
        SafeArrayPutElement(
            sa, 
            &idx as *const i32 as *const _, 
            &prop_id_val as *const i32 as *const _
        )?;

        // 注册属性监听
        // 注意：这里传入 sa 而不是 std::ptr::null()
        let result = automation.AddPropertyChangedEventHandler(
            &root,
            TreeScope_Descendants,
            None,
            &prop_interface,
            sa,
        );

        // 注册完成后，必须销毁 SAFEARRAY (因为它是 [in] 参数，我们负责创建和销毁)
        // 即使 AddProperty 失败也要销毁
        let _ = SafeArrayDestroy(sa);
        
        // 检查注册结果
        result?; 

        println!("  [+] 输入内容监听已注册");

        automation.AddAutomationEventHandler(
            UIA_Text_TextChangedEventId,
            &root,
            TreeScope_Descendants,
            None,
            &text_interface,
        )?;
        println!("  [+] TextChanged 监听已注册");
    }

    println!("\n正在运行... (按 Ctrl+C 结束)");
    let running = Arc::new(AtomicU32::new(1));
    while running.load(Ordering::Relaxed) > 0 {
        thread::sleep(Duration::from_millis(100));
    }

    unsafe {
        let _ = automation.RemoveAllEventHandlers();
        CoUninitialize();
    }
    Ok(())
}

// ==============================================================================
// 1. ManualFocusHandler (代码保持不变，为了完整性列出)
// ==============================================================================

#[repr(C)]
struct ManualFocusHandler {
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

// ==============================================================================
// 2. ManualPropertyHandler (修复了 VARIANT 处理)
// ==============================================================================

#[repr(C)]
struct ManualPropertyHandler {
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
        property_id: windows::Win32::UI::Accessibility::UIA_PROPERTY_ID,
        new_value: MaybeUninit<VARIANT>,
    ) -> HRESULT {
        // println!("    [属性变更] 属性ID: {:?}", property_id);
        if property_id == UIA_ValueValuePropertyId && !sender.is_null() {
            unsafe {
                let element: &IUIAutomationElement = std::mem::transmute(&sender);

                if let Ok(control_type) = element.CurrentControlType() {
                    if control_type == UIA_EditControlTypeId || control_type == UIA_GroupControlTypeId {
                        // 使用 assume_init_ref 避免不必要的复制，虽然 VARIANT 复制也不贵
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

fn get_text(element: &IUIAutomationElement) -> Result<String> {
    unsafe {
        // 1. 尝试 ValuePattern (适用于记事本、简单的 Win32 输入框)
        if let Ok(pattern_unk) = element.GetCurrentPattern(UIA_ValuePatternId) {
            // 必须检查是否为空指针 (有些应用返回 S_OK 但指针为 null)
            if let Ok(value_pattern) = pattern_unk.cast::<IUIAutomationValuePattern>() {
                    // 很多富文本控件虽然有 ValuePattern，但 CurrentValue 是空的，
                    // 所以这里获取到值后，如果非空才返回，否则继续尝试 TextPattern
                if let Ok(bstr) = value_pattern.CurrentValue() {
                    let s = bstr.to_string();
                    if !s.is_empty() {
                        return Ok(s);
                    }
                }
            }
        }

        // 2. 尝试 TextPattern (适用于飞书、Chrome、VS Code、Word)
        if let Ok(pattern_unk) = element.GetCurrentPattern(UIA_TextPatternId) {
            if let Ok(text_pattern) = pattern_unk.cast::<IUIAutomationTextPattern>() {
                // 获取整个文档范围
                if let Ok(range) = text_pattern.DocumentRange() {
                    // -1 表示获取没有长度限制的全部文本
                    if let Ok(bstr) = range.GetText(MAX_TEXT_LEN) {
                        return Ok(bstr.to_string());
                    }
                }
            }
        }
        
        // 3. 最后的兜底：尝试获取 Name 属性 (某些极端的自定义 UI 会把内容放在 Name 里)
        // 但通常不用这步，以免读出杂乱的标签名
        // element.CurrentName().map(|b| b.to_string())
        
        Ok(String::new())
    }
}

fn with_thread_automation<R>(f: impl FnOnce(&IUIAutomation) -> Result<R>) -> Result<R> {
    TL_AUTOMATION.with(|cell| {
        let mut automation = cell.borrow_mut();
        if automation.is_none() {
            unsafe {
                let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
            }
            let a: IUIAutomation = unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)? };
            *automation = Some(a);
        }
        f(automation.as_ref().unwrap())
    })
}

fn get_text_deep(element: &IUIAutomationElement) -> Result<String> {
    let text = get_text(element)?;
    if !text.is_empty() {
        return Ok(text);
    }

    with_thread_automation(|automation| unsafe {
        let cond_text: IUIAutomationCondition = automation.CreatePropertyCondition(
            UIA_IsTextPatternAvailablePropertyId,
            &VARIANT::from(true),
        )?;
        let cond_value: IUIAutomationCondition = automation.CreatePropertyCondition(
            UIA_IsValuePatternAvailablePropertyId,
            &VARIANT::from(true),
        )?;
        let cond: IUIAutomationCondition = automation.CreateOrCondition(&cond_text, &cond_value)?;

        let found = match element.FindFirst(TreeScope_Descendants, &cond) {
            Ok(found) => found,
            Err(_) => return Ok(String::new()),
        };

        get_text(&found)
    })
}

fn debounce_print(message: String) {
    let sender = DEBOUNCE_SENDER.get_or_init(|| {
        let (tx, rx) = mpsc::channel::<DebounceEvent>();
        thread::spawn(move || debounce_worker(rx));
        tx
    });
    let _ = sender.send(DebounceEvent { message });
}

fn debounce_worker(rx: mpsc::Receiver<DebounceEvent>) {
    loop {
        let mut last = match rx.recv() {
            Ok(ev) => ev,
            Err(_) => return,
        };
        loop {
            match rx.recv_timeout(Duration::from_millis(DEBOUNCE_MS)) {
                Ok(ev) => last = ev,
                Err(mpsc::RecvTimeoutError::Timeout) => {
                    println!("{}", last.message);
                    break;
                }
                Err(mpsc::RecvTimeoutError::Disconnected) => return,
            }
        }
    }
}

#[repr(C)]
struct ManualTextChangedHandler {
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
        event_id: windows::Win32::UI::Accessibility::UIA_EVENT_ID,
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
