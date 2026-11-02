# CorelDRAW Docker host integration

The host project supplies a COM-visible WinForms `UserControl` with a WPF `ElementHost` embedding `CardMakerPro.UI.Workspace`.

## What works at source level
- COM class identity and stable ProgID.
- WinForms/WPF hosting bridge.
- Dedicated host assembly and manual developer registration script.

## What is NOT implemented yet
- CorelDRAW-specific addon manifest / docker registration via `FrameWork.AddDocker`.
- CorelDRAW event subscriptions, document access, feature actions.
- Version detection and per-version adapter DLLs.
- CI build and MSI/EXE installer.
- Runtime testing inside X7 or any later release.

## Build requirements
- Windows machine with Visual Studio MSBuild, .NET Framework 4.6.2 developer targeting pack.
- Desktop Windows SDK assemblies, including WindowsFormsIntegration.
- CorelDRAW SDK matching the target version for the later registration adapter.

The docker control is a hosting scaffold, not an installable plugin.
