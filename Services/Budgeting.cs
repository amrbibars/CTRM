using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ClosedXML.Excel; // Ensure you have this for Export
using Microsoft.Win32;
using CTRM.Services;

namespace CTRM
{
    public partial class Budgeting : Window
    {
        private DataTable? _historicalData;
        private DataTable? _forecastData;
        private readonly string[] _months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

        public Budgeting()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Populate Start Month
            CmbStartMonth.ItemsSource = _months;
            CmbRollMonth.ItemsSource = _months;

            // Set defaults
            CmbStartMonth.SelectedIndex = 0;
            CmbRollMonth.SelectedIndex = 11; // Dec
        }

        // --- Import ---
        private void BtnImportOrders_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new() { Filter = "Excel Files|*.xlsx;*.xls;*.csv" };
                if (openFileDialog.ShowDialog() == true)
                {
                    _historicalData = ExcelReceiver.ImportOrders(openFileDialog.FileName);
                    OrdersGrid.ItemsSource = _historicalData.DefaultView;

                    // Auto-detect next likely year based on data
                    if (_historicalData.Rows.Count > 0)
                    {
                        var lastRow = _historicalData.Rows[^1];
                        int lastYear = Convert.ToInt32(lastRow["Year"]);
                        TxtRollYear.Text = (lastYear + 1).ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Import Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- Dynamic Controls Logic ---
        private void ChkNextYear_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool isChecked = ChkNextYear.IsChecked == true;
            TxtRollYear.IsEnabled = isChecked;

            // Refresh the RollTo list logic
            UpdateRollToOptions();
        }

        private void CmbStartMonth_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateRollToOptions();
        }

        private void UpdateRollToOptions()
        {
            if (CmbStartMonth.SelectedIndex == -1) return;

            // If Next Year is CHECKED: Roll To can be any month (Jan-Dec)
            if (ChkNextYear.IsChecked == true)
            {
                CmbRollMonth.ItemsSource = _months;
                if (CmbRollMonth.SelectedIndex == -1) CmbRollMonth.SelectedIndex = 11;
            }
            // If Next Year is UNCHECKED: Roll To must be > Start Month
            else
            {
                int startIdx = CmbStartMonth.SelectedIndex;
                string[] availableMonths;
                // Safety check: if Start is Dec, list is empty?
                if (startIdx == 11) availableMonths = ["Dec"]; // Fallback to allow single month forecast?
                else availableMonths = [.. _months.Skip(startIdx + 1)];

                CmbRollMonth.ItemsSource = availableMonths;
                CmbRollMonth.SelectedIndex = availableMonths.Length - 1; // Default to last available
            }
        }

        // Regex for Year Input (4 digits only)
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        // --- Forecast Execution ---
        private void BtnForecast_Click(object sender, RoutedEventArgs e)
        {
            if (_historicalData == null)
            {
                MessageBox.Show("Please import data first.");
                return;
            }

            // 1. Determine Start Date
            // We need to know the Year for the Start Month.
            // Logic: The Start Month usually follows immediately after the last historical data point.
            // Let's get the last date from history.
            DataRow lastHistRow = _historicalData.Rows[^1];
            int lastHistYear = Convert.ToInt32(lastHistRow["Year"]);
            string? lastHistMonthStr = lastHistRow["Month"].ToString();
            int lastHistMonth = Array.IndexOf(_months, lastHistMonthStr) + 1;

            int reqStartMonth = CmbStartMonth.SelectedIndex + 1;
            int reqStartYear = lastHistYear;

            // Simple logic: If requested start month <= last historical month, assume it's next year.
            // (Unless user manually sets years, but here we infer Start Year based on History end)
            if (reqStartMonth <= lastHistMonth)
            {
                reqStartYear = lastHistYear + 1;
            }

            // 2. Determine End Date
            int reqEndMonth;
            int reqEndYear;

            if (ChkNextYear.IsChecked == true)
            {
                if (string.IsNullOrEmpty(TxtRollYear.Text) || TxtRollYear.Text.Length != 4)
                {
                    MessageBox.Show("Please enter a valid 4-digit Target Year.");
                    return;
                }
                reqEndYear = int.Parse(TxtRollYear.Text);
                reqEndMonth = Array.IndexOf(_months, CmbRollMonth.SelectedItem.ToString()) + 1;
            }
            else
            {
                // Same year as Start Year
                reqEndYear = reqStartYear;
                // Since ItemsSource changes, SelectedItem is safer than Index
                if (CmbRollMonth.SelectedItem == null)
                    reqEndMonth = 12;
                else reqEndMonth = Array.IndexOf(_months, CmbRollMonth.SelectedItem.ToString()) + 1;
            }

            // Validation: End must be >= Start
            DateTime dStart = new(reqStartYear, reqStartMonth, 1);
            DateTime dEnd = new(reqEndYear, reqEndMonth, 1);

            if (dEnd < dStart)
            {
                MessageBox.Show("End date cannot be before Start date.");
                return;
            }

            // 3. Run Calculation
            BudgetingCalculator calc = new();
            _forecastData = calc.CalculateForecast(_historicalData, reqStartMonth, reqStartYear, reqEndMonth, reqEndYear);

            ForecastGrid.ItemsSource = _forecastData.DefaultView;
            BtnExport.IsEnabled = true;
        }

        // --- Export ---
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_forecastData == null) return;

            SaveFileDialog saveFileDialog = new() { Filter = "Excel Files|*.xlsx", FileName = "Forecast_Output.xlsx" };
            if (saveFileDialog.ShowDialog() == true)
            {
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("Forecast");
                    ws.Cell(1, 1).InsertTable(_forecastData);
                    wb.SaveAs(saveFileDialog.FileName);
                }
                MessageBox.Show("Export Successful!");
            }
        }

        private void DataGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            // 1. Check if the column name is "Year" to exclude it from formatting
            if (e.PropertyName == "Year")
            {
                return; // Do nothing for the Year column
            }

            // 2. Check if the property/column type is numeric for Market data
            if (e.PropertyType == typeof(double) || e.PropertyType == typeof(int) || e.PropertyType == typeof(decimal))
            {
                if (e.Column is DataGridTextColumn textColumn)
                {
                    // Apply the Comma Style format {0:N0} for orders/volumes
                    textColumn.Binding.StringFormat = "{0:N0}";

                    // Right-align market numbers for better professional readability
                    var style = new Style(typeof(TextBlock));
                    style.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Right));
                    style.Setters.Add(new Setter(TextBlock.MarginProperty, new Thickness(0, 0, 5, 0)));
                    textColumn.ElementStyle = style;
                }
            }
        }
    }
}