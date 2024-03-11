using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Interactivity;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Tmds.DBus.Protocol;

namespace PBI_Inventory.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // Getting country initials
            string countryInitialsPath = Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\..\PBI Inventory\Assets\country_initials.csv");
            Dictionary<string, string> CountryInitials = File.ReadLines(countryInitialsPath).Select(line => line.Split(',')).ToDictionary(line => line[0], line => line[1]);
            List<string> Countries = [.. CountryInitials.Keys];
            initialsAutocompleteList.ItemsSource = Countries;
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