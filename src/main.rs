mod debounce;
mod uia;

fn main() -> windows::core::Result<()> {
    uia::app::run()
}
