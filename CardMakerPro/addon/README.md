# Card Maker Pro CorelDRAW addon (X7-first test)

This add-on follows the CorelDRAW SDK XSLT mechanism. `AppUI.xslt` declares a `wpfhost` item and a `dockerData` entry named **Card Maker Pro**. `UserUI.xslt` declares user workspace integration; `CorelDrw.addon` marks the folder for CorelDRAW loading.

**Not yet runtime-tested.** Depending on CorelDRAW version/workspace, additional menu wiring may be required. The first test is whether Window > Dockers contains Card Maker Pro and can display the WPF Workspace.

The installer should copy this folder into `<CorelDRAW installation>\Programs64\Addons\CardMakerPro` (64-bit) or `Programs\Addons\CardMakerPro` (32-bit), alongside `CardMakerPro.UI.dll`. Do not install to the user's VBA macro directory. Existing macros must not be modified.

CorelDRAW official docs:
https://community.coreldraw.com/sdk/w/articles/173/custom-add-ons-in-coreldraw-corel-designer-and-corel-photo-paint
https://community.coreldraw.com/sdk/w/articles/174/custom-dockers-in-coreldraw-corel-designer-and-corel-photo-paint
