using System;
using System.Windows;
using Microsoft.Win32;
using CTRM.DB;
// using CTRM.Loaders; // Uncomment if IBV_Historical_Loader is in a different namespace

namespace CTRM.Services
{
    public partial class Forecasting : Window
    {
        public Forecasting()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Check if data was already loaded globally upon opening the window
            UpdateDataStatus();
        }
        private void btnLoadData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Title = "Select Historical Data (CSV Only)"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    // Assuming IBV_Historical_Loader is accessible in your project
                    var loader = new IBV_Historical_Loader();
                    var loadedData = loader.Load(openFileDialog.FileName);

                    if (loadedData != null && loadedData.Count > 0)
                    {
                        // 1. Send Data to the DB DataManager (Overwrites globally)
                        DataManager.Instance.SetData(loadedData);

                        MessageBox.Show("Data successfully loaded and updated across the application.",
                                        "Load Success", MessageBoxButton.OK, MessageBoxImage.Information);

                        // 2. Update UI to reflect new data
                        UpdateDataStatus();
                    }
                    else
                    {
                        MessageBox.Show("No valid Voice/Inbound data found.", "Load Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading CSV: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        private void UpdateDataStatus()
        {
            if (DataManager.Instance.HasData)
            {
                int recordCount = DataManager.Instance.AllData.Count;
                txtDataStatus.Text = $"Status: Data Loaded ({recordCount:N0} records active)";
                txtDataStatus.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#27AE60"));
            }
            else
            {
                txtDataStatus.Text = "Status: No data loaded. Please upload a file.";
                txtDataStatus.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E74C3C"));
            }
        }
    }
}