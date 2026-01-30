use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::Arc;
use std::thread;
use std::time::Duration;

use windows::core::Result;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize,
    CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED,
};
use windows::Win32::System::Variant::VT_I4;
use windows::Win32::System::Ole::{
    SafeArrayCreateVector, SafeArrayDestroy, SafeArrayPutElement,
};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationEventHandler,
    IUIAutomationFocusChangedEventHandler, IUIAutomationPropertyChangedEventHandler,
    TreeScope_Descendants, UIA_Text_TextChangedEventId, UIA_ValueValuePropertyId,
};

use crate::uia::handlers::{ManualFocusHandler, ManualPropertyHandler, ManualTextChangedHandler};

pub fn run() -> Result<()> {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).ok()?;
    }

    println!("初始化 UIA (手动实现 COM 模式)...");
    let automation: IUIAutomation = unsafe {
        CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER)?
    };

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
        automation.AddFocusChangedEventHandler(None, &focus_interface)?;
        println!("  [+] 焦点监听已注册");

        let root = automation.GetRootElement()?;

        let sa = SafeArrayCreateVector(VT_I4, 0, 1);

        if sa.is_null() {
            panic!("无法创建 SAFEARRAY: 内存不足");
        }

        let idx: i32 = 0;
        let prop_id_val = UIA_ValueValuePropertyId.0 as i32;

        SafeArrayPutElement(
            sa,
            &idx as *const i32 as *const _,
            &prop_id_val as *const i32 as *const _,
        )?;

        let result = automation.AddPropertyChangedEventHandler(
            &root,
            TreeScope_Descendants,
            None,
            &prop_interface,
            sa,
        );

        let _ = SafeArrayDestroy(sa);

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
