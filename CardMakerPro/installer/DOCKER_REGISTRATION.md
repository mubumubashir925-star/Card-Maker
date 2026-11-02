# CorelDRAW Docker registration: integration checklist

**Status: NOT REGISTERED.** COM registration alone does not expose a docker in CorelDRAW.

## Windows build host
- Build the solution with Visual Studio MSBuild on Windows and verify output assemblies.
- Use matching process bitness for the COM registration and CorelDRAW host.
- Register the COM-visible WinForms `CardMakerPro.DockerControl` class using the matching Framework RegAsm.
- Include the UI assembly and dependencies next to the host DLL.

## CorelDRAW-specific integration (next implementation)
- Select a specific CorelDRAW major version and inspect its shipped SDK AddDocker signature.
- Create a version-specific addon manifest and registration entry (not a generic registry guess).
- Call `Application.FrameWork.AddDocker` with the SDK-defined parameters and stable Docker GUID.
- Test that the docker is available from CorelDRAW's Window > Dockers menu, opens, closes, and survives restart.
- Repeat against X7 and modern versions; verify process architecture and OS compatibility.
- Add cleanup/unregistration and safe upgrade logic to the installer.

**Do not run register-host.cmd expecting a visible CorelDRAW docker.** The registration script registers only the COM class. No user-facing installer is ready.
