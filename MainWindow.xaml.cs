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
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            _ = DownloadLatestSetupAsync(); // Using the _ = ... syntax to call the async method without awaiting it directly.
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
            string setupUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20148.exe";
            string localPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "setup.exe");

            using (HttpClient client = new HttpClient())
            {
                var response = await client.GetAsync(setupUrl);

                if (response.IsSuccessStatusCode)
                {
                    var bytes = await response.Content.ReadAsByteArrayAsync();
                    await File.WriteAllBytesAsync(localPath, bytes);
                }
                else
                {
                    MessageBox.Show("An Error occured");
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

            string setupPath = @"path\to\setup.exe";  // Modify this to the actual path of the ODT setup.exe
            string arguments = $"/configure \"{XmlPathTextBox.Text}\"";

            Process.Start(setupPath, arguments);
        }
    }
}
