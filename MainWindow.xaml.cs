using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace OfficeInstaller
{
    public partial class MainWindow : Window
    {
        private const string SetupUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20148.exe";
        private readonly string _localSetupPath;
        private static readonly HttpClient _httpClient = new HttpClient();

        public MainWindow()
        {
            InitializeComponent();

            _localSetupPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "setup.exe");
            _ = DownloadLatestSetupAsync();
            PopulateXmlFiles();
        }

        private void ConfigureXmlButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://config.office.com/deploymentsettings");
        }


        private void PopulateXmlFiles()
        {
            var xmlFiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.xml")
                .Select(file => new XmlFile(Path.GetFileName(file), file))
                .ToList();

            XmlFilesComboBox.Items.Add(XmlFile.Placeholder);
            xmlFiles.ForEach(xml => XmlFilesComboBox.Items.Add(xml));
            XmlFilesComboBox.SelectedIndex = 0;
        }

        private async Task DownloadLatestSetupAsync()
        {
            if (File.Exists(_localSetupPath)) return;

            try
            {
                byte[] setupData = await _httpClient.GetByteArrayAsync(SetupUrl);
                await File.WriteAllBytesAsync(_localSetupPath, setupData);

                MessageBox.Show("Setup was successfully downloaded.");
            }
            catch (HttpRequestException)
            {
                MessageBox.Show("Failed to download the setup. Please check your internet connection.");
            }
            catch (IOException)
            {
                MessageBox.Show("Failed to write the setup file to disk. Ensure you have sufficient permissions and disk space.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An unexpected error occurred: {ex.Message}");
            }
        }

        private void XmlFilesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (XmlFilesComboBox.SelectedItem is XmlFile selectedXmlFile)
            {
                XmlPathTextBox.Text = selectedXmlFile != XmlFile.Placeholder ? selectedXmlFile.FullPath : string.Empty;
            }
        }

        private void XmlPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ValidateXml(XmlPathTextBox.Text);
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
                ValidateXml(XmlPathTextBox.Text);
            }
        }

        private void ValidateXml(string path)
        {
            if (XmlUtilities.IsValidXml(path))
            {
                ValidationCheckmark.Visibility = Visibility.Visible;
            }
            else
            {
                ValidationCheckmark.Visibility = Visibility.Collapsed;
                MessageBox.Show("The chosen XML is not valid.");
            }
        }

        private void InstallOfficeButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(XmlPathTextBox.Text) || XmlPathTextBox.Text == XmlFile.Placeholder.FullPath)
            {
                MessageBox.Show("Please select an XML configuration file first.");
                return;
            }

            try
            {
                Process.Start(_localSetupPath, $"/configure \"{XmlPathTextBox.Text}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Installation failed: {ex.Message}");
            }
        }

        private void ShowExeFolderButton_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("explorer.exe", AppDomain.CurrentDomain.BaseDirectory);
        }
    }



    public class XmlFile
    {
        public string FileName { get; }
        public string FullPath { get; }

        public XmlFile(string fileName, string fullPath)
        {
            FileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
            FullPath = fullPath ?? throw new ArgumentNullException(nameof(fullPath));
        }

        public static XmlFile Placeholder { get; } = new XmlFile("Select a Configuration XML...", "");

        public override string ToString() => FileName;
    }


    public static class XmlUtilities
    {
        public static bool IsValidXml(string path)
        {
            try
            {
                var xmlDocument = new System.Xml.XmlDocument();
                xmlDocument.Load(path);
                return true; // If XML is successfully loaded
            }
            catch
            {
                return false; // Any error in parsing will result in an invalid XML
            }
        }
    }
}
