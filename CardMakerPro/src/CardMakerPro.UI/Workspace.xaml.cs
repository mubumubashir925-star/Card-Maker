using System.Windows;
using System.Windows.Controls;
using System.Reflection;
using System.Threading.Tasks;
using System.Diagnostics;

namespace CardMakerPro.UI
{
    public partial class Workspace : UserControl
    {
        public Workspace()
        {
            InitializeComponent();
        }

        private async void CheckUpdates_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null) button.IsEnabled = false;
            StatusText.Text = "Checking GitHub Releases...";
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                var release = await Task.Run(() => ReleaseChecker.Check(version));
                StatusText.Text = release.IsNewer ? "Update available: " + release.Version : "Latest release: " + release.Version;
                if (release.IsNewer && MessageBox.Show("Version " + release.Version + " is available. Open the official GitHub release page?", "Card Maker Pro", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    Process.Start(release.ReleaseUrl);
            }
            catch (System.Exception ex)
            {
                StatusText.Text = "Update check unavailable";
                MessageBox.Show(ex.Message, "Update check", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally { if (button != null) button.IsEnabled = true; }
        }

        private void Navigate_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;
            var page = button.Tag as string ?? "Cards";
            PageTitle.Text = page;
            PageDescription.Text = page == "QR / Barcode"
                ? "QR and barcode generator • color, rotation and PowerClip"
                : page + " workspace";
            StatusText.Text = "UI preview • " + page + " • Host adapter not connected";
        }
    }
}
