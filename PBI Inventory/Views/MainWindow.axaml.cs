using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Tmds.DBus.Protocol;

namespace PBI_Inventory.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void ExportClickHandler(object sender, RoutedEventArgs args)
        {
            if (message.Text == "Export done !")
            {
                message.Text = "Ready to export...";
                return;
            }
            message.Text = "Export done !";
        }
    }
}