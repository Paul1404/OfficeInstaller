using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;
using Microsoft.Win32;
using System.Net.Http;
using System.IO;


namespace OfficeInstaller
{
    public partial class MainWindow : Window
    {
        private const string SetupUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20148.exe";
        private readonly string _localSetupPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "setup.exe");

        public MainWindow()
        {
            InitializeComponent();
            _ = DownloadLatestSetupAsync();
        }

        private void ChooseXmlButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                XmlPathTextBox.Text = openFileDialog.FileName;
            }
        }

        private async Task DownloadLatestSetupAsync()
        {
            using (var client = new HttpClient())
            {
                try
                {
                    var response = await client.GetAsync(SetupUrl);
                    response.EnsureSuccessStatusCode();

                    var bytes = await response.Content.ReadAsByteArrayAsync();
                    await File.WriteAllBytesAsync(_localSetupPath, bytes);
                }
                catch (Exception ex)  // Catching a general exception for demonstration. It's a good practice to catch more specific exceptions.
                {
                    MessageBox.Show($"An error occurred: {ex.Message}");
                }
            }
        }

        private void InstallOfficeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(XmlPathTextBox.Text))
            {
                MessageBox.Show("Please select an XML configuration file first.");
                return;
            }

            string arguments = $"/configure \"{XmlPathTextBox.Text}\"";

            try
            {
                Process.Start(_localSetupPath, arguments);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Installation failed: {ex.Message}");
            }
        }
    }
}
