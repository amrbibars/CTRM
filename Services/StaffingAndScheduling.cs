using CTRM.Models;
using CTRM.Services;
using ClosedXML.Excel; // Ensure this is referenced for the upgraded Export function
using Microsoft.Win32;
using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Data;
using System.Windows.Media;
using System.IO;
using System.Linq;

namespace CTRM
{
    public partial class StaffingAndScheduling : Window
    {
        // ---------------------------------------------------------
        // 1) VARIABLES & FIELDS
        // ---------------------------------------------------------

        private WeeklyMatrix? _volumes;
        private WeeklyMatrix? _aht;
        private enum StaffingConcept { None, ErlangC, Workload, Mixed }

        // ---------------------------------------------------------
        // 2) REGEX GENERATORS (New)
        // ---------------------------------------------------------

        // Generated Regex for Decimals (e.g., 0.85, .5, 10)
        private static readonly Regex RegexDecimalPattern = new(@"^[0-9]*([.][0-9]*)?$", RegexOptions.Compiled);

        // Generated Regex for Integers (e.g., 100, 3600)
        private static readonly Regex RegexIntegerPattern = new(@"^[0-9]+$", RegexOptions.Compiled);

        // ---------------------------------------------------------
        // 3) CONSTRUCTOR
        // ---------------------------------------------------------

        public StaffingAndScheduling()
        {
            InitializeComponent();
            PreviewGrid.AutoGeneratingColumn += PreviewGrid_AutoGeneratingColumn;
            PreviewGrid.LoadingRow += PreviewGrid_LoadingRow;

            ConceptErlangC.Checked += Concept_Checked;
            ConceptWorkload.Checked += Concept_Checked;
            ConceptMixed.Checked += Concept_Checked;
        }


        // ---------------------------------------------------------
        // 3) BUTTON CLICK HANDLERS (same order as XAML)
        // ---------------------------------------------------------

        // 3.1 Load Volumes
        private void LoadVolumes_Click(object sender, RoutedEventArgs e)
        {
            var path = PickExcel();
            if (path is null) return;

            try
            {
                _volumes = ExcelReceiver.LoadMatrix(path);
                MessageBox.Show($"Loaded volumes: {_volumes.Rows.Count} intervals, {_volumes.Headers.Count} columns.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load volumes: " + ex.Message);
            }
            BtnExport.IsEnabled = false;
        }

        // 3.2 Load AHT
        private void LoadAht_Click(object sender, RoutedEventArgs e)
        {
            var path = PickExcel();
            if (path is null) return;

            try
            {
                _aht = ExcelReceiver.LoadMatrix(path);
                MessageBox.Show($"Loaded AHT: {_aht.Rows.Count} intervals, {_aht.Headers.Count} columns.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to load AHT: " + ex.Message);
            }
            BtnExport.IsEnabled = false;
        }

        // 3.3 Calculate Staffing
        private void CalculateStaffing_Click(object sender, RoutedEventArgs e)
        {
            var concept = GetSelectedConcept();

            if (concept == StaffingConcept.None)
            {
                MessageBox.Show("Please select a staffing concept first.");
                return;
            }

            if (!EnsureDataLoaded()) return;

            // Get inputs including new Shift and Days Off fields
            if (!TryGetInputs(concept, out float sla, out int timeSec, out float occ,
                              out float shrink, out int intervalSec, out float shiftDur, out int daysOff))
            {
                return;
            }

                switch (concept)
            {
                case StaffingConcept.ErlangC:
                    CalculateErlangC(sla, timeSec, occ, shrink, intervalSec, shiftDur, daysOff);
                    break;

                case StaffingConcept.Workload:
                    CalculateWorkload(occ, shrink, intervalSec, shiftDur, daysOff);
                    break;

                case StaffingConcept.Mixed:
                    CalculateMixed(sla, timeSec, occ, shrink, intervalSec, shiftDur, daysOff);
                    break;
            }
        }

        // ---------------------------------------------------------
        // 4) CALCULATION FUNCTIONS
        // ---------------------------------------------------------

        private void CalculateErlangC(float sla, int timeSec, float occ, float shrink, int intervalSec, float shiftDur, int daysOff)
        {
            if (_volumes is null || _aht is null) return;

            int rowCount = _volumes.Rows.Count;
            int colCount = _volumes.Headers.Count;
            if (_volumes.HasTotalsRow && rowCount > 0) rowCount --;

            DataTable dt = new ("StaffingResults");
            dt.Columns.Add("Interval", typeof(string));
            for (int c = 0; c < colCount; c++) dt.Columns.Add(_volumes.Headers[c], typeof(object));

            double[] dayNetSums = new double[colCount];

            for (int r = 0; r < rowCount; r++)
            {
                var row = dt.NewRow();
                var start = TimeSpan.Zero;
                var labelTime = start + TimeSpan.FromSeconds(intervalSec * r);
                row["Interval"] = labelTime.ToString(@"hh\:mm");

                for (int c = 0; c < colCount; c++)
                {
                    double calls = _volumes.Rows[r].Values[c];
                    double aht = _aht.Rows[r].Values[c];

                    var (net, finalAgents) = ComputeErlangForCell(calls, aht, sla, timeSec, occ, shrink, intervalSec);
                    row[c + 1] = finalAgents;
                    dayNetSums[c] += net;
                }
                dt.Rows.Add(row);
            }

            AddTotalRow(dt, dayNetSums, intervalSec, shiftDur, daysOff, shrink);
            PreviewGrid.ItemsSource = dt.DefaultView;
            BtnExport.IsEnabled = true;
        }

        private void CalculateWorkload(float occ, float shrink, int intervalSec, float shiftDur, int daysOff)
        {
            if (_volumes is null || _aht is null) return;

            int rowCount = _volumes.Rows.Count;
            int colCount = _volumes.Headers.Count;
            if (_volumes.HasTotalsRow && rowCount > 0) rowCount --;

            DataTable dt = new ("StaffingResults");
            dt.Columns.Add("Interval", typeof(string));
            for (int c = 0; c < colCount; c++) dt.Columns.Add(_volumes.Headers[c], typeof(object));

            double[] dayNetSums = new double[colCount];

            for (int r = 0; r < rowCount; r++)
            {
                var row = dt.NewRow();
                row["Interval"] = TimeSpan.FromSeconds(intervalSec * r).ToString(@"hh\:mm");

                for (int c = 0; c < colCount; c++)
                {
                    double calls = _volumes.Rows[r].Values[c];
                    double aht = _aht.Rows[r].Values[c];

                    // Step 1: Net Agents = (Volume * AHT) / (Interval * Occupancy)
                    double numerator = calls * aht;
                    double denominator = intervalSec * occ;
                    int net = (denominator == 0) ? 0 : (int)Math.Ceiling(numerator / denominator);

                    int buffer = (int)Math.Round(net * shrink, MidpointRounding.AwayFromZero);
                    row[c + 1] = net + buffer;
                    dayNetSums[c] += net;
                }
                dt.Rows.Add(row);
            }

            AddTotalRow(dt, dayNetSums, intervalSec, shiftDur, daysOff, shrink);
            PreviewGrid.ItemsSource = dt.DefaultView;
            BtnExport.IsEnabled = true;
        }

        private void CalculateMixed(float sla, int timeSec, float occ, float shrink, int intervalSec, float shiftDur, int daysOff)
        {
            MessageBox.Show("Mixed fallback to Erlang C.");
            CalculateErlangC(sla, timeSec, occ, shrink, intervalSec, shiftDur, daysOff);
        }

        // New Helper: Adds the Total Row using the sequential multiplier logic
        private static void AddTotalRow(DataTable dt, double[] dayNetSums, int intervalSec, float shiftDur, int daysOff, float shrink)
        {
            var totalRow = dt.NewRow();
            totalRow["Interval"] = "Total";
            for (int c = 0; c < dayNetSums.Length; c++)
            {
                // Step 1: Divide by 2 if 1800, else 1
                double workloadHours = dayNetSums[c] / (intervalSec == 1800 ? 2.0 : 1.0);
                // Step 2: Base Heads
                double baseHeads = workloadHours / shiftDur;
                // Step 3: Sequential Days Off
                double afterDaysOff = baseHeads * (1.0 + (daysOff / 7.0));
                // Step 4: Sequential Shrinkage
                double finalWithShrink = afterDaysOff * (1.0 + shrink);
                // Step 5: Round to natural number
                totalRow[c + 1] = (int)Math.Round(finalWithShrink, MidpointRounding.AwayFromZero);
            }
            dt.Rows.Add(totalRow);
        }

        // ---------------------------------------------------------
        // 5) INPUT PARSING AND VALIDATION
        // ---------------------------------------------------------

        private StaffingConcept GetSelectedConcept()
        {
            if (ConceptErlangC.IsChecked == true) return StaffingConcept.ErlangC;
            if (ConceptWorkload.IsChecked == true) return StaffingConcept.Workload;
            if (ConceptMixed.IsChecked == true) return StaffingConcept.Mixed;
            return StaffingConcept.None;
        }

        // UI function to make inputs inactive for Workload concept
        private void Concept_Checked(object sender, RoutedEventArgs e)
        {
            if (TxtSla == null || TxtTimeSec == null) return;
            bool isWorkload = ConceptWorkload.IsChecked == true;
            TxtSla.IsEnabled = !isWorkload;
            TxtTimeSec.IsEnabled = !isWorkload;
        }

        private int GetIntervalSeconds()
        {
            if (Interval1800.IsChecked == true) return 1800;
            if (Interval3600.IsChecked == true) return 3600;
            return 3600;
        }

        private bool TryGetInputs(StaffingConcept concept, out float sla, out int timeSec, out float occ,
                                  out float shrink, out int intervalSec, out float shiftDur, out int daysOff)
        {
            sla = occ = shrink = shiftDur = 0f;
            timeSec = intervalSec = daysOff = 0;

            // SLA & Time (Skip validation for Workload)
            if (concept != StaffingConcept.Workload)
            {
                if (!float.TryParse(TxtSla.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out sla) || sla < 0 || sla > 1)
                { MessageBox.Show("SLA must be 0–1."); return false; }
                if (!int.TryParse(TxtTimeSec.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out timeSec) || timeSec <= 0)
                { MessageBox.Show("Time must be positive."); return false; }
            }

            if (!float.TryParse(TxtOcc.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out occ) || occ <= 0 || occ > 1)
            { MessageBox.Show("Occupancy must be 0–1."); return false; }

            if (!float.TryParse(TxtShrink.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out shrink) || shrink < 0 || shrink > 1)
            { MessageBox.Show("Shrinkage must be 0–1."); return false; }

            if (!float.TryParse(TxtShiftDuration.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out shiftDur) || shiftDur <= 0)
            { MessageBox.Show("Net Shift Duration must be positive."); return false; }

            if (!int.TryParse(TxtDaysOff.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out daysOff) || daysOff < 0)
            { MessageBox.Show("Days Off must be >= 0."); return false; }

            intervalSec = GetIntervalSeconds();
            return true;
        }



        private bool EnsureDataLoaded()
        {
            if (_volumes is null || _aht is null)
            {
                MessageBox.Show("Please load both Volumes and AHT first.");
                return false;
            }
            return true;
        }

        // ... (PickExcel and Filter Handlers remain as provided in your version) ...

        private static string? PickExcel()
        {
            var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb|All Files|*.*", Title = "Select Excel file" };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        private void Decimal_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckRegex(sender, e, RegexDecimalPattern);
        }

        private void Int_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckRegex(sender, e, RegexIntegerPattern);
        }

        private void Decimal_Paste(object sender, DataObjectPastingEventArgs e)
        {
            CheckPaste(e, RegexDecimalPattern);
        }

        private void Int_Paste(object sender, DataObjectPastingEventArgs e)
        {
            CheckPaste(e, RegexIntegerPattern);
        }

        // Helper to check text input against a Regex
        private static void CheckRegex(object sender, TextCompositionEventArgs e, Regex r)
        {
            var tb = (TextBox)sender;
            string candidate = tb.Text.Remove(tb.SelectionStart, tb.SelectionLength)
                                 .Insert(tb.SelectionStart, e.Text);
            e.Handled = !r.IsMatch(candidate);
        }

        // Helper to check paste operations
        private static void CheckPaste(DataObjectPastingEventArgs e, Regex r)
        {
            if (e.DataObject.GetDataPresent(DataFormats.Text))
            {
                string text = (string)e.DataObject.GetData(DataFormats.Text)!;
                if (!r.IsMatch(text)) e.CancelCommand();
            }
            else
            {
                e.CancelCommand();
            }
        }

        // ----------------------------
        // VISUAL LOGIC
        // ----------------------------

        private void PreviewGrid_AutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == "Interval") { e.Cancel = true; } // Hide Interval column in body
            else { e.Column.MinWidth = 80; }
        }

        private void PreviewGrid_LoadingRow(object? sender, DataGridRowEventArgs e)
        {
            if (e.Row.Item is DataRowView rowView)
            {
                e.Row.Header = rowView["Interval"].ToString();
                if (rowView["Interval"].ToString() == "Total")
                {
                    e.Row.FontWeight = FontWeights.Bold;
                    e.Row.Background = new SolidColorBrush(Color.FromRgb(0xEF, 0xFF, 0xD9));
                }

                else
                {
                    e.Row.FontWeight = FontWeights.Normal;
                    e.Row.Background = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                }

            }
            }

        // ----------------------------
        // Erlang C HELPER FUNCTIONS
        // ----------------------------

        private static double ComputeANFactor(double A, int N)
        {
            double result = 1.0;
            for (int i = 1; i <= N; i++) result *= A / i;
            return result;
        }

        private static double ErlangC(int N, double A)
        {
            if (A >= N) return 1.0;
            double sum = 0.0;
            for (int k = 0; k < N; k++)
            {
                double term = 1.0;
                for (int i = 1; i <= k; i++) term *= A / i;
                sum += term;
            }
            double an_fact = 1.0;
            for (int i = 1; i <= N; i++) an_fact *= A / i;
            double num = an_fact * (N / (N - A));
            return num / (sum + num);
        }

        private static (int netAgents, int finalAgents) ComputeErlangForCell(double calls, double aht, float sla, int timeSec, float occ, float shrink, int intervalSec)
        {
            if (calls <= 0 || aht <= 0) return (0, 0);
            double A = calls * aht / intervalSec;
            if (A <= 0) return (0, 0);
            int N = Math.Max(1, (int)Math.Round(A));
            while (A / N > occ) N++;
            int bestN = N;
            for (int i = 0; i < 500; i++)
            {
                double Pw = ErlangC(bestN, A);
                double slaQueued = 1.0 - (Pw * Math.Exp((A - bestN) * (timeSec / aht)));
                if (slaQueued >= sla) break;
                bestN++;
            }
            int buffer = (int)Math.Round(bestN * shrink, MidpointRounding.AwayFromZero);
            return (bestN, bestN + buffer);
        }

        // ----------------------------
        // EXPORT
        // ----------------------------

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (PreviewGrid.ItemsSource is not DataView view) return;
            var dlg = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", FileName = "StaffingResults.xlsx" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Staffing");
                var dt = view.ToTable();
                var tableRange = ws.Cell(1, 1).InsertTable(dt);
                tableRange.Theme = XLTableTheme.TableStyleMedium2;
                ws.Columns().AdjustToContents();
                wb.SaveAs(dlg.FileName);
                MessageBox.Show("Export successful.");
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }
    }
}