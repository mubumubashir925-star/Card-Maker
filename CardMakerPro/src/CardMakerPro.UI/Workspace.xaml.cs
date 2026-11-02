using System.Windows;
using System.Windows.Controls;

namespace CardMakerPro.UI
{
    public partial class Workspace : UserControl
    {
        public Workspace()
        {
            InitializeComponent();
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
