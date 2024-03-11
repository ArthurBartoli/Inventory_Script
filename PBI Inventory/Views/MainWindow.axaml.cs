using Avalonia;
using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Interactivity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Tmds.DBus.Protocol;

namespace PBI_Inventory.Views
{
    public partial class MainWindow : Window

    {
        private Dictionary<string, string> CountryInitialsDict { get; set; }    
        private string CountryInitials { get; set; }
        private bool exportOption { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            // Initialize boolean
            exportOption = false;

            // Giving country initials to autocomplete 
            string countryInitialsPath = Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\..\PBI Inventory\Assets\country_initials.csv");
            CountryInitialsDict = File.ReadLines(countryInitialsPath).Select(line => line.Split(',')).ToDictionary(line => line[0], line => line[1]);
            List<string> Countries = [.. CountryInitialsDict.Keys];
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

        public void CountryInitialsSelection(object sender, SelectionChangedEventArgs args)
        {
            try
            {
                string selection = (string) args.AddedItems[0];
                countrySelected.Text = CountryInitialsDict[selection];
                CountryInitials = selection;
            }
            catch (System.ArgumentOutOfRangeException) {
                countrySelected.Text = "No country selected";
            }

        }
        public void onExportOptionChecked(object sender, RoutedEventArgs args)
        {
            if (exportOptionText.Text == "The export option is not checked") {
                exportOption = true;
                exportOptionText.Text = "The export option is checked !";
                return;
            }
            exportOption = false;
            exportOptionText.Text = "The export option is not checked";
        }
    }
}    