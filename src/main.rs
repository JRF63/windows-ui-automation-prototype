use windows::Win32::{
    Foundation::POINT,
    System::Com::{CoCreateInstance, CoInitialize, CLSCTX_INPROC_SERVER},
    UI::{
        Accessibility::{
            CUIAutomation8, IUIAutomation, IUIAutomationCondition, IUIAutomationElement,
        },
        WindowsAndMessaging::GetCursorPos,
    },
};

use std::{mem::MaybeUninit, thread, time::Duration};

/// Helper for `GetCursorPos`.
fn get_cursor_pos() -> windows::core::Result<POINT> {
    // Or init with zeroes using `POINT::default` instead of `MaybeUninit`
    let mut point = MaybeUninit::<POINT>::uninit();
    unsafe {
        GetCursorPos(point.as_mut_ptr())?;
        Ok(point.assume_init())
    }
}

fn find_text_pattern_element(
    ui_automation: &IUIAutomation,
    element: &IUIAutomationElement,
) -> windows::core::Result<IUIAutomationElement> {
    todo!()
}

fn main() -> windows::core::Result<()> {
    let sleep_dur_for_polling = Duration::from_millis(200);

    // Windows syscalls are unsafe
    unsafe {
        CoInitialize(None)?;

        // COM usage is in the same process
        let class_context = CLSCTX_INPROC_SERVER;
        let ui_automation: IUIAutomation = CoCreateInstance(&CUIAutomation8, None, class_context)?;

        loop {
            let point = get_cursor_pos()?;
            match ui_automation.ElementFromPoint(point) {
                Ok(element) => {
                    break;
                }
                Err(_) => thread::sleep(sleep_dur_for_polling),
            }
        }
    }

    Ok(())
}
