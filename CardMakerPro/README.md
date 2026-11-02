# Card Maker Pro — CorelDRAW Plugin

A modern, dockable CorelDRAW ID card production workspace. This directory is the **new plugin project**, separate from the proven `Card_Maker.frm` / `Card_Maker.frx` VBA source.

## Target
- CorelDRAW Graphics Suite X7 through current Windows Graphics Suite releases.
- Windows 7 for X7 (legacy build), Windows 10/11 for supported modern releases.
- WPF user interface, CorelDRAW .NET docker hosting.
- Per-version adapters and build configurations; do **not** assume binary compatibility.
- No automatic modification of the existing VBA macro.

## Development milestones
1. Bootstrap WPF docker host, version detection, and a per-version CorelDRAW SDK reference strategy.
2. Implement the complete professional workspace: Cards, Data, Images, QR/Barcode, Arrange, Settings.
3. Port working VBA operations feature by feature, preserving the existing macro as the behavior reference.
4. Native CorelDRAW color dialog where available, HEX/RGB/CMYK input, eyedropper support where exposed by SDK.
5. PowerClip alignment modes (center, top-center, fill, fit), correct aspect ratio, Code 128 quiet zones.
6. Barcode rotation: auto, 0, 90, 180, 270 degrees, plus preview.
7. Per-version release artifacts and a version-aware installer/uninstaller.

## Architecture
- `src/CardMakerPro.Contracts`: host-independent interfaces and models.
- `src/CardMakerPro.UI`: WPF views and view models.
- `src/CardMakerPro.Corel`: CorelDRAW API adapters, built against installed SDK for each target.
- `installer`: installation manifest and future installer scripts.

## Important
This is an **architecture/bootstrap scaffold**, not yet a compiled or installable plugin. Do not deploy it to production CorelDRAW until a matching host adapter, docker registration, and installer are built.

## Official SDK references
- X7 Docker API: https://community.coreldraw.com/sdk/api/draw/17/m/framework.adddocker
- 2025 Docker API: https://community.coreldraw.com/sdk/api/draw/26/m/framework.adddocker
- Addon guide: https://community.coreldraw.com/sdk/w/articles
