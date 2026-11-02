# Card Maker — UI redesign (02 November 2026)

This dated experimental copy is isolated from the production `Card_Maker.frm` and `Card_Maker.frx` files.

## Compatibility
CorelDRAW X7 / Windows 7 and later, MSForms VBA.

## Changes
- Removed tiny/unreliable Unicode glyphs that rendered poorly on Windows 7.
- Normalized action labels to short, readable text.
- Retained existing VBA event dispatch and tooltip descriptions.
- Existing card generation and PowerClip logic is unchanged.

## Install/test
1. Back up your current GMS project.
2. Download this folder's `Card_Maker.frm` and the original root `Card_Maker.frx` into **the same local directory**.
3. Import the form into a test VBA project. The form uses the FRX resource, so the FRX must be next to the FRM.
4. Test all actions before replacing the working project.

## Important limitation
This is an interim readable design, **not** the requested finished image-based icon UI. Real BMP icons require a separate tested MSForms picture-loading implementation and assets. Do not merge into main as a finished visual redesign.
