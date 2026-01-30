use std::sync::mpsc::{self};
use std::thread;
use std::time::Duration;

use crate::global::{DEBOUNCE_MS, DEBOUNCE_SENDER, DebounceEvent};

pub fn debounce_print(message: String) {
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
