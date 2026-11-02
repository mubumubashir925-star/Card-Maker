using System;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace CardMakerPro.UI
{
    public sealed class ReleaseInfo
    {
        public string Version { get; set; }
        public string ReleaseUrl { get; set; }
        public string Notes { get; set; }
        public bool IsNewer { get; set; }
    }

    public static class ReleaseChecker
    {
        private const string LatestUrl = "https://api.github.com/repos/mubumubashir925-star/Card-Maker/releases/latest";
        public static ReleaseInfo Check(string installedVersion)
        {
            using (var client = new WebClient())
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                client.Headers.Add(HttpRequestHeader.UserAgent, "CardMakerPro-Updater");
                client.Headers.Add(HttpRequestHeader.Accept, "application/vnd.github+json");
                var json = client.DownloadString(LatestUrl);
                var data = new JavaScriptSerializer().DeserializeObject(json) as System.Collections.Generic.Dictionary<string, object>;
                if (data == null) throw new InvalidOperationException("Invalid GitHub release response.");
                var tag = Convert.ToString(data["tag_name"]);
                var url = Convert.ToString(data["html_url"]);
                var notes = data.ContainsKey("body") ? Convert.ToString(data["body"]) : "";
                if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps ||
                    !string.Equals(uri.Host, "github.com", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Invalid release URL.");
                Version latest, installed;
                bool newer = Version.TryParse(tag.TrimStart('v', 'V'), out latest)
                    && Version.TryParse(installedVersion.TrimStart('v', 'V'), out installed)
                    && latest > installed;
                return new ReleaseInfo { Version = tag, ReleaseUrl = url, Notes = notes, IsNewer = newer };
            }
        }
    }
}
