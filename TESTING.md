# Card Maker V2 — implementation workspace

- `old-version/`: snapshot of the previous FRM/FRX pair. Keep both files together.
- `new-version/`: runtime VBA form with two-column Card Maker tools, QR/barcode advanced controls, RGB colors and numbered position status.
- `ui-preview/`: conceptual UI SVG (not a real CorelDRAW screenshot).
- `designer-setup/`: **optional, experimental** VBA macro for rearranging existing *designer* controls in CorelDRAW X7.

## Designer UI limitation

The current `new-version/Card_Maker.frx` is still the legacy designer binary. The modern runtime controls are created in code and do not automatically appear in the VBA editor. To make the editor itself show a modern layout, the actual MSForms designer must be edited and re-exported. `designer-setup/ApplyModernCardMakerDesigner.bas` attempts a first-pass arrangement of existing controls through VBIDE; it does **not** produce the complete proposed mockup or recreate runtime-only controls. It has not been tested in CorelDRAW X7.

## Suggested test procedure

1. Back up `GlobalMacros.gms` and work in a test GMS.
2. Import `new-version/Card_Maker.frm` with `Card_Maker.frx` in the same folder.
3. Compile the VBA project and test generation with two sample records.
4. For editor layout experiments, import the `.bas` module and run `ApplyModernCardMakerDesigner`. This requires VBA project model access.
5. Inspect positions, save, export the designer-updated FRM and FRX as a new pair, then test again.

Do not use this branch for production until these steps pass. No real CorelDRAW runtime test has been performed.
