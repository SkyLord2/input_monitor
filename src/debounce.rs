use std::sync::mpsc::{self, Sender};
use std::sync::OnceLock;
use std::thread;
use std::time::Duration;

const DEBOUNCE_MS: u64 = 200;

struct DebounceEvent {
    message: String,
}

static DEBOUNCE_SENDER: OnceLock<Sender<DebounceEvent>> = OnceLock::new();

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
