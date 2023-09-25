mod focus;

use windows::Win32::{
    System::{
        Com::{CoCreateInstance, CoInitialize, CLSCTX_INPROC_SERVER},
        Variant::VARIANT,
    },
    UI::Accessibility::{
        CUIAutomation8, IUIAutomation, IUIAutomationCondition, IUIAutomationElement,
        IUIAutomationTextPattern2, IUIAutomationValuePattern, TreeScope_Subtree,
        UIA_IsTextPattern2AvailablePropertyId, UIA_IsValuePatternAvailablePropertyId,
        UIA_TextPattern2Id, UIA_ValueIsReadOnlyPropertyId, UIA_ValuePatternId, UIA_PROPERTY_ID,
    },
};

use std::{mem::ManuallyDrop, thread, time::Duration};

/// Helper for instantiating a boolean `windows::Win32::System::Variant::VARIANT`.
fn create_bool_variant(value: bool) -> VARIANT {
    // Putting this here to avoid polluting the namespace
    use windows::Win32::{
        Foundation::{VARIANT_FALSE, VARIANT_TRUE},
        System::Variant::{VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_BOOL},
    };

    VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                vt: VT_BOOL,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: VARIANT_0_0_0 {
                    boolVal: if value { VARIANT_TRUE } else { VARIANT_FALSE },
                },
            }),
        },
    }
}

fn create_condition(
    ui_automation: &IUIAutomation,
    conditions: &[(UIA_PROPERTY_ID, bool)],
) -> windows::core::Result<IUIAutomationCondition> {
    unsafe {
        let mut res = ui_automation.CreateTrueCondition()?;
        for (prop_id, t) in conditions {
            let additional =
                ui_automation.CreatePropertyCondition(*prop_id, create_bool_variant(*t))?;
            res = ui_automation.CreateAndCondition(&res, &additional)?;
        }
        Ok(res)
    }
}

fn create_text_avail_condition(
    ui_automation: &IUIAutomation,
) -> windows::core::Result<IUIAutomationCondition> {
    create_condition(
        ui_automation,
        &[(UIA_IsTextPattern2AvailablePropertyId, true)],
    )
}

fn create_editable_text_avail_condition(
    ui_automation: &IUIAutomation,
) -> windows::core::Result<IUIAutomationCondition> {
    create_condition(
        ui_automation,
        &[
            (UIA_IsTextPattern2AvailablePropertyId, true),
            (UIA_ValueIsReadOnlyPropertyId, false),
        ],
    )
}

fn find_first_element_in_subtree(
    element: &IUIAutomationElement,
    condition: &IUIAutomationCondition,
) -> windows::core::Result<Option<IUIAutomationElement>> {
    unsafe {
        match element.FindFirst(TreeScope_Subtree, condition) {
            Ok(m) => Ok(Some(m)),
            Err(e) => {
                // Returns an `Error` with `HRESULT(0x00000000)` (i.e. success) if nothing satisfies the
                // condition
                if e.code().is_ok() {
                    Ok(None)
                } else {
                    Err(e)
                }
            }
        }
    }
}

fn main() -> windows::core::Result<()> {
    let polling_delay = Duration::from_secs(2);
    let sleep_dur_for_polling = Duration::from_millis(200);

    unsafe {
        CoInitialize(None)?;

        // COM usage is in the same process
        let class_context = CLSCTX_INPROC_SERVER;
        let ui_automation: IUIAutomation = CoCreateInstance(&CUIAutomation8, None, class_context)?;

        let condition = create_condition(&ui_automation, &[(UIA_IsTextPattern2AvailablePropertyId, true)])?;
        // let condition = create_condition(
        //     &ui_automation,
        //     &[
        //         (UIA_IsTextPattern2AvailablePropertyId, true),
        //         (UIA_ValueIsReadOnlyPropertyId, false),
        //     ],
        // )?;
        // let condition = create_condition(
        //     &ui_automation,
        //     &[
        //         (UIA_IsValuePatternAvailablePropertyId, true),
        //         (UIA_IsTextPattern2AvailablePropertyId, true),
        //         (UIA_ValueIsReadOnlyPropertyId, false),
        //     ],
        // )?;

        thread::sleep(polling_delay);
        loop {
            if let Ok(parent) = focus::get_focused_element(&ui_automation) {
                if let Some(e) = find_first_element_in_subtree(&parent, &condition)? {
                    // println!(
                    //     "Parent: {}, Control type: {}",
                    //     parent.CurrentControlType()?.0,
                    //     e.CurrentControlType()?.0
                    // );

                    let text_pattern: IUIAutomationTextPattern2 =
                        e.GetCurrentPatternAs(UIA_TextPattern2Id)?;

                    let selection = text_pattern.GetSelection()?;

                    for i in 0..selection.Length()? {
                        let ee = selection.GetElement(i)?;
                        println!("- {}\n {}", i, ee.GetText(1_000_000)?);
                    }

                    let mut b = windows::Win32::Foundation::TRUE;
                    let mut r = std::mem::MaybeUninit::uninit();

                    if let Ok(_) = text_pattern.GetCaretRange(&mut b, r.as_mut_ptr()) {
                        if let Some(range) = r.assume_init() {
                            let text = range.GetText(1_000_000)?;
                            println!("- Caret\n {}", text);
                        }
                    }
                }
            }

            thread::sleep(sleep_dur_for_polling);
        }
    }
}
