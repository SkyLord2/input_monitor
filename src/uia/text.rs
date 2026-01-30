use std::cell::RefCell;

use windows::core::{Interface, Result, VARIANT};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx,
    CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED,
};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationCondition, IUIAutomationElement,
    IUIAutomationTextPattern, IUIAutomationValuePattern, TreeScope_Descendants,
    UIA_IsTextPatternAvailablePropertyId, UIA_IsValuePatternAvailablePropertyId,
    UIA_TextPatternId, UIA_ValuePatternId,
};

thread_local! {
    static TL_AUTOMATION: RefCell<Option<IUIAutomation>> = const { RefCell::new(None) };
}

const MAX_TEXT_LEN: i32 = 4096;

pub fn get_text_deep(element: &IUIAutomationElement) -> Result<String> {
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

fn get_text(element: &IUIAutomationElement) -> Result<String> {
    unsafe {
        if let Ok(pattern_unk) = element.GetCurrentPattern(UIA_ValuePatternId) {
            if let Ok(value_pattern) = pattern_unk.cast::<IUIAutomationValuePattern>() {
                if let Ok(bstr) = value_pattern.CurrentValue() {
                    let s = bstr.to_string();
                    if !s.is_empty() {
                        return Ok(s);
                    }
                }
            }
        }

        if let Ok(pattern_unk) = element.GetCurrentPattern(UIA_TextPatternId) {
            if let Ok(text_pattern) = pattern_unk.cast::<IUIAutomationTextPattern>() {
                if let Ok(range) = text_pattern.DocumentRange() {
                    if let Ok(bstr) = range.GetText(MAX_TEXT_LEN) {
                        return Ok(bstr.to_string());
                    }
                }
            }
        }

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
