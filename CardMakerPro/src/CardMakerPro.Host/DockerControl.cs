using System;
using System.Runtime.InteropServices;
using System.Windows.Forms.Integration;
using System.Runtime.InteropServices;
using CardMakerPro.UI;

namespace CardMakerPro.Host
{
    // WinForms container exposes an HWND to CorelDRAW's native docker host.
    // WPF runs inside ElementHost; no CorelDRAW PIA is required to compile this UI layer.
    [ComVisible(true)]
    [Guid("F46B2D41-831D-47C8-95C5-B4CC7C0A01A2")]
    [ProgId("CardMakerPro.DockerControl")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class DockerControl : System.Windows.Forms.UserControl
    {
        private readonly ElementHost _host;

        public DockerControl()
        {
            Dock = System.Windows.Forms.DockStyle.Fill;
            _host = new ElementHost();
            _host.Dock = System.Windows.Forms.DockStyle.Fill;
            _host.Child = new Workspace();
            Controls.Add(_host);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _host.Dispose();
            base.Dispose(disposing);
        }
    }
}
