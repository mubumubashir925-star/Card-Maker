# First CorelDRAW Docker smoke test

**Experimental. Not a production installer.** CorelDRAW's documented `FrameWork.AddDocker(Guid, ClassName, AssemblyPath)` accepts a .NET class and its DLL path, but whether our WinForms/WPF host is accepted by every version must be tested.

1. Open the successful GitHub Actions run and download `CardMakerPro-Host-Libraries`.
2. Extract the ZIP to a permanent folder, e.g. `C:\CardMakerPro-Test\`. Ensure `CardMakerPro.Host.dll` and `CardMakerPro.UI.dll` are both available together. If the artifact stores them in separate directories, copy them together before testing.
3. Back up the CorelDRAW workspace and your existing VBA/GMS macro files. Do not overwrite existing macros.
4. Start CorelDRAW X7 with a blank test document. Open VBA editor with Alt+F11.
5. File > Import File, choose `CorelDockerBootstrap.bas` and import it into a temporary macro project.
6. Run `CardMakerProBootstrap.RegisterCardMakerProDocker`.
7. Enter the full path to `CardMakerPro.Host.dll`.
8. Check whether the Card Maker Pro docker appears. Capture the complete error text if the API rejects the class or assembly.

## Important limits
- The script makes an AddDocker API call; it does not install permanent menu entries or an addon manifest.
- The DLLs have been built on CI but not tested in CorelDRAW. COM registration is not assumed to be necessary for AddDocker with an assembly path.
- Do not reset the CorelDRAW workspace with F8 without backing it up.
- X7/Windows 7 requires checking that the targeted .NET Framework runtime is supported and installed.
- A CorelDRAW addon package with `Coreldrw.addon`, `AppUI.xslt` and `UserUI.xslt` will be built after the initial docker host is verified.
