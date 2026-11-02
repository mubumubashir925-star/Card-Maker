# Card Maker Pro Preview Setup

This is a **developer-preview setup**, not a working CorelDRAW add-on yet.

The installer:
- Copies the compiled WPF host and UI assemblies into a stable Windows installation folder.
- Includes the VBA test bootstrap and documentation for developers.
- Provides a normal Windows uninstaller.

It does **not**:
- Automatically register a Docker in the CorelDRAW menu.
- Integrate Card Maker VBA features.
- Provide automatic in-app update installation.
- Guarantee X7/Windows 7 compatibility.

The missing permanent addon manifest and Docker UI registration must be implemented and tested before releasing a production one-click installer. Do not distribute this preview as a finished plugin.

GitHub Actions builds the setup using Inno Setup on Windows. The resulting file is named `CardMakerPro-Preview-Setup.exe`.
