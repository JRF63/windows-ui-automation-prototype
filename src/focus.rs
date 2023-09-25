use windows::Win32::{
    Foundation::POINT,
    UI::{
        Accessibility::{IUIAutomation, IUIAutomationElement},
        WindowsAndMessaging::GetCursorPos,
    },
};

use std::mem::MaybeUninit;

/// Helper for `GetCursorPos`.
fn get_cursor_pos() -> windows::core::Result<POINT> {
    // Or init with zeroes using `POINT::default` instead of `MaybeUninit`
    let mut point = MaybeUninit::<POINT>::uninit();
    unsafe {
        GetCursorPos(point.as_mut_ptr())?;
        Ok(point.assume_init())
    }
}

/// Returns the element at the cursor.
pub fn get_cursor_element(
    ui_automation: &IUIAutomation,
) -> windows::core::Result<IUIAutomationElement> {
    let point = get_cursor_pos()?;
    unsafe { ui_automation.ElementFromPoint(point) }
}

/// Returns the element focused by the keyboard.
pub fn get_focused_element(
    ui_automation: &IUIAutomation,
) -> windows::core::Result<IUIAutomationElement> {
    unsafe { ui_automation.GetFocusedElement() }
}