use std::cell::RefCell;
use std::sync::mpsc::Sender;
use std::sync::OnceLock;

use windows::Win32::UI::Accessibility::IUIAutomation;

pub(crate) const MAX_TEXT_LEN: i32 = 4096;
pub(crate) const DEBOUNCE_MS: u64 = 200;

pub(crate) struct DebounceEvent {
    pub(crate) message: String,
}

pub(crate) static DEBOUNCE_SENDER: OnceLock<Sender<DebounceEvent>> = OnceLock::new();

thread_local! {
    pub(crate) static TL_AUTOMATION: RefCell<Option<IUIAutomation>> = const { RefCell::new(None) };
}
