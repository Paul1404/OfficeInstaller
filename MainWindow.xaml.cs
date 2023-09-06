using Microsoft.Win32;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using System;
using System.Linq;

namespace OfficeInstaller
{
    public partial class MainWindow : Window
    {
        private const string SetupUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20148.exe";
        private readonly string _localSetupPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "setup.exe");

        public List<string> XmlFiles { get; set; } = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            _ = DownloadLatestSetupAsync();
            PopulateXmlFiles();
        }

        public class XmlFile
        {
            public string FileName { get; set; }
            public string FullPath { get; set; }

            public XmlFile(string fileName, string fullPath)
            {
                FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
                FullPath = fullPath ?? throw new ArgumentNullException(nameof(fullPath));
            }

            public static XmlFile Placeholder => new XmlFile("Select a Configuration XML...", "");

            public override string ToString()
            {
                return FileName;
            }
        }


        private void PopulateXmlFiles()
        {
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string[] xmlFiles = Directory.GetFiles(appDirectory, "*.xml");

            // Ensure the placeholder is the first item.
            if (XmlFilesComboBox.Items.Count == 0 || (XmlFilesComboBox.Items[0] as XmlFile)?.FileName != XmlFile.Placeholder.FileName)
            {
                XmlFilesComboBox.Items.Insert(0, XmlFile.Placeholder);
            }

            // Add the XML files.
            foreach (var xmlFile in xmlFiles)
            {
                XmlFilesComboBox.Items.Add(new XmlFile(Path.GetFileName(xmlFile), xmlFile));
            }

            // Set the default item as selected.
            XmlFilesComboBox.SelectedIndex = 0;
        }




        private void XmlFilesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (XmlFilesComboBox.SelectedItem is XmlFile selectedXmlFile)
            {
                if (selectedXmlFile == XmlFile.Placeholder)
                {
                    XmlPathTextBox.Text = ""; // Reset/clear the textbox.
                }
                else
                {
                    XmlPathTextBox.Text = selectedXmlFile.FullPath;
                }
            }
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
            using var client = new HttpClient();
            try
            {
                var response = await client.GetAsync(SetupUrl);
                response.EnsureSuccessStatusCode();

                var bytes = await response.Content.ReadAsByteArrayAsync();
                await File.WriteAllBytesAsync(_localSetupPath, bytes);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }

        private void InstallOfficeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(XmlPathTextBox.Text) || XmlPathTextBox.Text == "Select an XML file...")
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

        private void ShowExeFolderButton_Click(object sender, RoutedEventArgs e)
        {
            string exeFolderPath = AppDomain.CurrentDomain.BaseDirectory;
            Process.Start("explorer.exe", exeFolderPath);
        }

    }
}
