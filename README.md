# Card Maker — separated versions

## old-version/
Snapshot of the files on main when this rebuild branch was started. Keep both `.frm` and `.frx` together. This snapshot includes the previous runtime redesign; the repository's earlier history contains the initial designer.

## new-version/
Work-in-progress layout for CorelDRAW X7. Card data and batch PowerClip / proportional image fit / stretch fit / text fit are arranged in two columns; QR and barcode are an optional advanced area. Position status is shown as numbered SET / NOT SET labels.

**Important:** The `.frx` resource is still a copy of the existing designer resource. The VBA editor will still show the legacy design. A true editor-visible redesign requires rebuilding native controls in the UserForm designer and exporting the resulting FRM/FRX pair from CorelDRAW. Do not represent the concept preview as a tested screenshot.

## ui-preview/
Design concept only, not a screenshot from CorelDRAW.

Do not overwrite the working GMS without a backup. Compile and test with a sample document before production.
