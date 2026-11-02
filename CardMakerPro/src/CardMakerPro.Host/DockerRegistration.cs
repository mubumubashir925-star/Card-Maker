using System;
using System.Reflection;

namespace CardMakerPro.Host
{
    public static class DockerRegistration
    {
        // Host-specific CorelDRAW SDK registration is intentionally separate.
        // The AddDocker invocation signature and add-on XML must be verified
        // against each installed CorelDRAW version before deployment.
        public const string ControlProgId = "CardMakerPro.DockerControl";
        public const string DockerGuid = "F46B2D41-831D-47C8-95C5-B4CC7C0A01A2";

        public static string VerifyControl()
        {
            var type = Type.GetTypeFromProgID(ControlProgId, false);
            return type == null
                ? "Docker COM control is not registered."
                : "Docker COM control registered: " + type.FullName;
        }

        public static string HostAssemblyVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }
    }
}
