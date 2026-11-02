# Installation design (not yet an installer)

## Target detection
1. Discover installed CorelDRAW Graphics Suite editions and executable architecture.
2. Read exact major version and choose a matching signed build; refuse unsupported editions.
3. Check required .NET Framework runtime and CorelDRAW automation support.
4. Back up existing plugin configuration; **never overwrite** `Card_Maker.gms`.
5. Install the matching add-on assets and register the docker according to that version's SDK.
6. Restart CorelDRAW and open the Card Maker Pro docker from its registered menu/toolbar.
7. Provide uninstall and rollback.

## Compatibility matrix (planned, not validated)
| CorelDRAW | Host generation | Build |
| --- | --- | --- |
| X7 / 17.x | Legacy .NET Framework | x86/x64 SDK-specific |
| X8 / 18.x–2020 / 22.x | Legacy/intermediate | SDK-specific |
| 2021–2026 | Modern Windows releases | SDK-specific |

## Installer implementation
A Windows installer (Inno Setup or WiX) will be introduced after host registration and first compiled build are verified. No downloadable installer is claimed yet.

## User prerequisites
CorelDRAW Graphics Suite edition with automation/add-on support. CorelDRAW Standard and other limited editions may not expose these APIs.
