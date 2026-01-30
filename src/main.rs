mod debounce;
mod global;
mod uia;

fn main() -> windows::core::Result<()> {
    uia::app::run()
}
