using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using ClosedXML.Excel;
using CTRM.DB; // <-- CRITICAL: Ensure you have this using statement

namespace CTRM
{
    // ---------------------------------------------------------
    // 1. DATA STRUCTURE
    // ---------------------------------------------------------
    public struct IBV_CallData
    {
        public DateTime Date;
        public TimeSpan StartTime;
        public string Team;
        public int Offered;
        public int Aht;
        public int WeekNumber;
        public int Year;
    }

    public class WeeklyPivotRow
    {
        public string WeekLabel { get; set; } = string.Empty;
        public int Year { get; set; }
        public int WeekNumber { get; set; }
        public int Sun { get; set; }
        public int Mon { get; set; }
        public int Tue { get; set; }
        public int Wed { get; set; }
        public int Thu { get; set; }
        public int Fri { get; set; }
        public int Sat { get; set; }
        public int WeekTotal => Sun + Mon + Tue + Wed + Thu + Fri + Sat;
    }

    public class ArrivalIntervalRow
    {
        public TimeSpan Time { get; set; }
        public string CustomLabel { get; set; }
        public string Label => !string.IsNullOrEmpty(CustomLabel) ? CustomLabel : Time.ToString(@"hh\:mm");
        public double SunVal { get; set; }
        public double MonVal { get; set; }
        public double TueVal { get; set; }
        public double WedVal { get; set; }
        public double ThuVal { get; set; }
        public double FriVal { get; set; }
        public double SatVal { get; set; }
        public List<double> Values => new List<double> { SunVal, MonVal, TueVal, WedVal, ThuVal, FriVal, SatVal };
        public string SunDisplay { get; set; } = "";
        public string MonDisplay { get; set; } = "";
        public string TueDisplay { get; set; } = "";
        public string WedDisplay { get; set; } = "";
        public string ThuDisplay { get; set; } = "";
        public string FriDisplay { get; set; } = "";
        public string SatDisplay { get; set; } = "";
        public Brush SunColor { get; set; } = Brushes.White;
        public Brush MonColor { get; set; } = Brushes.White;
        public Brush TueColor { get; set; } = Brushes.White;
        public Brush WedColor { get; set; } = Brushes.White;
        public Brush ThuColor { get; set; } = Brushes.White;
        public Brush FriColor { get; set; } = Brushes.White;
        public Brush SatColor { get; set; } = Brushes.White;
    }

    // ---------------------------------------------------------
    // 2. LOGIC ENGINE
    // ---------------------------------------------------------
    public class IBV_Historical_Loader
    {
        public List<IBV_CallData> Load(string filePath)
        {
            var results = new List<IBV_CallData>(10000);
            using (var reader = new StreamReader(filePath))
            {
                string? headerLine = reader.ReadLine();
                if (string.IsNullOrEmpty(headerLine)) return results;

                headerLine = headerLine.Replace("\uFEFF", "").Replace("\"", "").Trim();
                var headers = headerLine.Split(',').Select(h => h.Trim().ToUpper()).ToList();

                int idxDate = headers.IndexOf("LABEL_YYYY_MM_DD");
                int idxInterval = headers.IndexOf("LABEL_YYYY_MM_DD_HH24_30INT");
                int idxOffered = headers.IndexOf("OFFERED");
                int idxTeam = headers.IndexOf("RESOURCE_NAME");
                int idxMedia = headers.IndexOf("MEDIA_NAME");
                int idxType = headers.IndexOf("INTERACTION_TYPE");

                if (idxDate == -1 || idxInterval == -1 || idxOffered == -1)
                {
                    string found = string.Join(" | ", headers);
                    throw new Exception($"Column Mismatch!\n\nExpected: LABEL_YYYY_MM_DD, OFFERED, etc.\n\nFound in file: {found}");
                }

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var columns = line.Split(',');

                    if (idxMedia != -1 && idxMedia < columns.Length && columns[idxMedia].Trim() != "Voice") continue;
                    if (idxType != -1 && idxType < columns.Length && columns[idxType].Trim() != "Inbound") continue;

                    try
                    {
                        if (!DateTime.TryParseExact(columns[idxDate].Trim().Replace("\"", ""),
                            new[] { "M/d/yyyy", "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy" },
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                            continue;

                        string rawInterval = columns[idxInterval];
                        if (!rawInterval.Contains(" ")) continue;

                        string timePart = rawInterval.Split(' ')[1];
                        TimeSpan startTime = TimeSpan.Parse(timePart.Split('-')[0]);

                        int offered = (int)double.Parse(columns[idxOffered]);
                        string team = (idxTeam != -1 && idxTeam < columns.Length) ? columns[idxTeam].Trim() : "Unknown";

                        var weekInfo = GetFiscalWeek(date);

                        results.Add(new IBV_CallData
                        {
                            Date = date,
                            StartTime = startTime,
                            Team = team,
                            Offered = offered,
                            Year = weekInfo.Year,
                            WeekNumber = weekInfo.Week
                        });
                    }
                    catch { continue; }
                }
            }
            return results;
        }

        private (int Year, int Week) GetFiscalWeek(DateTime date)
        {
            DateTime startOfWeek = date.AddDays(-(int)date.DayOfWeek);
            DateTime endOfWeek = startOfWeek.AddDays(6);
            int fiscalYear = startOfWeek.Year;

            if (endOfWeek.Year > startOfWeek.Year) return (endOfWeek.Year, 1);

            DateTime jan1 = new DateTime(fiscalYear, 1, 1);
            DateTime week1Start = jan1.AddDays(-(int)jan1.DayOfWeek);
            int daysDiff = (startOfWeek - week1Start).Days;
            int weekNum = (daysDiff / 7) + 1;

            return (fiscalYear, weekNum);
        }
    }

    // ---------------------------------------------------------
    // 3. MAIN WINDOW LOGIC
    // ---------------------------------------------------------
    public partial class Arrival_Pattern : Window
    {
        // Notice we removed _allData, as we will use DataManager.Instance.AllData
        private List<IBV_CallData> _filteredData;
        private bool _isDataLoading = false;

        public Arrival_Pattern()
        {
            InitializeComponent();
            _filteredData = new List<IBV_CallData>();
        }

        private void LoadData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            openFileDialog.Title = "Select Historical Data (CSV Only)";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var loader = new IBV_Historical_Loader();
                    var loadedData = loader.Load(openFileDialog.FileName);

                    if (loadedData != null && loadedData.Count > 0)
                    {
                        _isDataLoading = true;

                        // 1. Send Data to the DB DataManager
                        DataManager.Instance.SetData(loadedData);

                        // Use data from the DataManager for setup
                        var masterData = DataManager.Instance.AllData;

                        // 2. Populate SKILLS ListBox
                        var teams = masterData.Select(x => x.Team).Distinct().OrderBy(t => t).ToList();
                        teams.Insert(0, "All");
                        lstTeams.ItemsSource = teams;
                        if (teams.Count > 0)
                        {
                            lstTeams.SelectedIndex = 0;
                            btnSkillSelect.Content = "All Skills Selected";
                        }

                        // 3. Populate WEEKS ListBox
                        var weeks = masterData.Select(x => x.WeekNumber.ToString()).Distinct().OrderBy(w => w.Length).ThenBy(w => w).ToList();
                        weeks.Insert(0, "All Weeks (Total)");
                        lstWeeks.ItemsSource = weeks;

                        if (weeks.Count > 0)
                        {
                            lstWeeks.SelectedIndex = 0;
                            btnWeekSelect.Content = "All Weeks (Total)";
                        }

                        _isDataLoading = false;
                        RefreshDashboard();
                    }
                    else
                    {
                        MessageBox.Show("No valid Voice/Inbound data found.", "Load Error");
                    }
                }
                catch (Exception ex)
                {
                    _isDataLoading = false;
                    MessageBox.Show("Error reading CSV: " + ex.Message);
                }
            }
        }

        private void OnFilterChanged(object sender, RoutedEventArgs e)
        {
            if (_isDataLoading) return;

            if (sender == lstTeams)
            {
                var count = lstTeams.SelectedItems.Count;
                if (count == 0) return;

                var first = lstTeams.SelectedItems[0].ToString();
                if (first == "All")
                    btnSkillSelect.Content = "All Skills Selected";
                else
                    btnSkillSelect.Content = count == 1 ? first : $"{count} Skills Selected";
            }
            RefreshDashboard();
        }

        private void OnIntervalFilterChanged(object sender, RoutedEventArgs e)
        {
            if (_isDataLoading) return;

            if (sender == lstWeeks)
            {
                var count = lstWeeks.SelectedItems.Count;
                if (count == 0) return;

                var first = lstWeeks.SelectedItems[0].ToString();
                if (first.Contains("Total"))
                    btnWeekSelect.Content = "All Weeks (Total)";
                else
                    btnWeekSelect.Content = count == 1 ? $"Week {first}" : $"{count} Weeks Selected";
            }
            GenerateIntervalGrid();
        }

        private void RefreshDashboard()
        {
            // Connect to DataManager instead of local list
            var masterData = DataManager.Instance.AllData;
            if (masterData == null || masterData.Count == 0) return;

            var selectedItems = lstTeams.SelectedItems.Cast<string>().ToList();
            if (selectedItems.Count == 0) return;

            if (selectedItems.Contains("All"))
                _filteredData = masterData;
            else
                _filteredData = masterData.Where(x => selectedItems.Contains(x.Team)).ToList();

            var pivotRows = _filteredData
                .GroupBy(x => new { x.Year, x.WeekNumber })
                .OrderBy(g => g.Key.Year).ThenBy(g => g.Key.WeekNumber)
                .Select(g => new WeeklyPivotRow
                {
                    WeekLabel = $"W{g.Key.WeekNumber}",
                    Year = g.Key.Year,
                    WeekNumber = g.Key.WeekNumber,
                    Sun = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Sunday).Sum(x => x.Offered),
                    Mon = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Monday).Sum(x => x.Offered),
                    Tue = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Tuesday).Sum(x => x.Offered),
                    Wed = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Wednesday).Sum(x => x.Offered),
                    Thu = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Thursday).Sum(x => x.Offered),
                    Fri = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Friday).Sum(x => x.Offered),
                    Sat = g.Where(x => x.Date.DayOfWeek == DayOfWeek.Saturday).Sum(x => x.Offered),
                }).ToList();

            gridHistory.ItemsSource = pivotRows;
            DrawChart(pivotRows);
            GenerateIntervalGrid();
        }

        private void GenerateIntervalGrid()
        {
            if (_filteredData == null || _filteredData.Count == 0) return;
            if (lstWeeks.SelectedItems.Count == 0) return;

            var selectedItems = lstWeeks.SelectedItems.Cast<string>().ToList();
            List<IBV_CallData> scopeData;

            if (selectedItems.Contains("All Weeks (Total)"))
            {
                scopeData = _filteredData;
            }
            else
            {
                var selectedWeekNums = selectedItems.Select(int.Parse).ToList();
                scopeData = _filteredData.Where(x => selectedWeekNums.Contains(x.WeekNumber)).ToList();
            }

            var intervalRows = new List<ArrivalIntervalRow>();
            double[,] matrix = new double[48, 7];
            double[] daySums = new double[7];

            for (int i = 0; i < 48; i++)
            {
                TimeSpan ts = TimeSpan.FromMinutes(i * 30);
                var newRow = new ArrivalIntervalRow { Time = ts };
                intervalRows.Add(newRow);

                for (int d = 0; d < 7; d++)
                {
                    DayOfWeek day = (DayOfWeek)d;
                    var matching = scopeData.Where(x => x.StartTime == ts && x.Date.DayOfWeek == day).ToList();

                    double val = matching.Sum(x => x.Offered);

                    matrix[i, d] = val;
                    daySums[d] += val;

                    switch (day)
                    {
                        case DayOfWeek.Sunday: newRow.SunVal = val; break;
                        case DayOfWeek.Monday: newRow.MonVal = val; break;
                        case DayOfWeek.Tuesday: newRow.TueVal = val; break;
                        case DayOfWeek.Wednesday: newRow.WedVal = val; break;
                        case DayOfWeek.Thursday: newRow.ThuVal = val; break;
                        case DayOfWeek.Friday: newRow.FriVal = val; break;
                        case DayOfWeek.Saturday: newRow.SatVal = val; break;
                    }
                }
            }

            double minVal = double.MaxValue, maxVal = 0;
            bool isPercent = rbPattern.IsChecked == true;

            for (int i = 0; i < 48; i++)
            {
                for (int d = 0; d < 7; d++)
                {
                    double raw = matrix[i, d];
                    double finalVal = raw;
                    if (isPercent && daySums[d] > 0) finalVal = raw / daySums[d];

                    if (finalVal > maxVal) maxVal = finalVal;
                    if (finalVal < minVal) minVal = finalVal;
                }
            }

            for (int i = 0; i < 48; i++)
            {
                var row = intervalRows[i];
                for (int d = 0; d < 7; d++)
                {
                    double raw = matrix[i, d];
                    double finalVal = raw;
                    if (isPercent && daySums[d] > 0) finalVal = raw / daySums[d];

                    string text = isPercent ? finalVal.ToString("P1") : raw.ToString("N0");
                    Brush color = GetHeatmapColor(finalVal, minVal, maxVal);

                    switch ((DayOfWeek)d)
                    {
                        case DayOfWeek.Sunday: row.SunDisplay = text; row.SunColor = color; break;
                        case DayOfWeek.Monday: row.MonDisplay = text; row.MonColor = color; break;
                        case DayOfWeek.Tuesday: row.TueDisplay = text; row.TueColor = color; break;
                        case DayOfWeek.Wednesday: row.WedDisplay = text; row.WedColor = color; break;
                        case DayOfWeek.Thursday: row.ThuDisplay = text; row.ThuColor = color; break;
                        case DayOfWeek.Friday: row.FriDisplay = text; row.FriColor = color; break;
                        case DayOfWeek.Saturday: row.SatDisplay = text; row.SatColor = color; break;
                    }
                }
            }

            if (!isPercent)
            {
                var totalRow = new ArrivalIntervalRow { CustomLabel = "TOTAL" };
                totalRow.SunDisplay = daySums[0].ToString("N0");
                totalRow.MonDisplay = daySums[1].ToString("N0");
                totalRow.TueDisplay = daySums[2].ToString("N0");
                totalRow.WedDisplay = daySums[3].ToString("N0");
                totalRow.ThuDisplay = daySums[4].ToString("N0");
                totalRow.FriDisplay = daySums[5].ToString("N0");
                totalRow.SatDisplay = daySums[6].ToString("N0");

                totalRow.SunVal = daySums[0];
                totalRow.MonVal = daySums[1];
                totalRow.TueVal = daySums[2];
                totalRow.WedVal = daySums[3];
                totalRow.ThuVal = daySums[4];
                totalRow.FriVal = daySums[5];
                totalRow.SatVal = daySums[6];
                totalRow.SunColor = totalRow.MonColor = totalRow.TueColor = totalRow.WedColor =
                totalRow.ThuColor = totalRow.FriColor = totalRow.SatColor = Brushes.LightGray;
                intervalRows.Add(totalRow);
            }

            gridIntervals.ItemsSource = intervalRows;
        }

        private Brush GetHeatmapColor(double value, double min, double max)
        {
            if (max <= min) return new SolidColorBrush(Color.FromRgb(245, 245, 245));

            double t = (value - min) / (max - min);
            byte r, g, b;

            if (t < 0.5)
            {
                double localT = t * 2;
                r = (byte)(200 + (255 - 200) * localT);
                g = (byte)(225 + (255 - 225) * localT);
                b = 255;
            }
            else
            {
                double localT = (t - 0.5) * 2;
                r = 255;
                g = (byte)(255 - (255 - 215) * localT);
                b = (byte)(255 - (255 - 215) * localT);
            }

            return new SolidColorBrush(Color.FromRgb(r, g, b));
        }

        private void OnChartSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (gridHistory.ItemsSource is List<WeeklyPivotRow> data)
            {
                DrawChart(data);
            }
        }

        private void DrawChart(List<WeeklyPivotRow> data)
        {
            chartCanvas.Children.Clear();
            if (data == null || data.Count == 0) return;

            double w = chartCanvas.ActualWidth;
            double h = chartCanvas.ActualHeight;
            if (w < 10 || h < 10) return;

            string viewMode = (cmbView.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Weekly Total";
            List<double> vals = new List<double>();

            if (viewMode == "All Days (Trend)")
            {
                foreach (var r in data)
                {
                    vals.AddRange(new double[] { r.Sun, r.Mon, r.Tue, r.Wed, r.Thu, r.Fri, r.Sat });
                }
            }
            else if (viewMode == "Weekly Total")
            {
                vals = data.Select(x => (double)x.WeekTotal).ToList();
            }
            else
            {
                if (Enum.TryParse(viewMode, out DayOfWeek target))
                {
                    vals = data.Select(r => target switch
                    {
                        DayOfWeek.Monday => (double)r.Mon,
                        DayOfWeek.Tuesday => (double)r.Tue,
                        DayOfWeek.Wednesday => (double)r.Wed,
                        DayOfWeek.Thursday => (double)r.Thu,
                        DayOfWeek.Friday => (double)r.Fri,
                        DayOfWeek.Saturday => (double)r.Sat,
                        _ => (double)r.Sun
                    }).ToList();
                }
            }

            if (vals.Count < 2) return;
            double max = vals.Max(); if (max == 0) max = 1;

            Polyline line = new Polyline { Stroke = Brushes.DodgerBlue, StrokeThickness = 2 };
            double stepX = w / (vals.Count - 1);

            for (int i = 0; i < vals.Count; i++)
            {
                double x = i * stepX;
                double y = h - (vals[i] / max * h);
                line.Points.Add(new Point(x, y));

                Ellipse dot = new Ellipse { Fill = Brushes.White, Stroke = Brushes.DodgerBlue, StrokeThickness = 1, Width = 6, Height = 6 };
                Canvas.SetLeft(dot, x - 3); Canvas.SetTop(dot, y - 3);
                chartCanvas.Children.Add(dot);
            }
            chartCanvas.Children.Add(line);
        }

        // =========================================================
        // EXPORT LOGIC (EXCEL .XLSX)
        // =========================================================

        private string GetSelectedItemsString(ListBox listBox)
        {
            if (listBox.SelectedItems.Count == 0) return "None";
            var items = listBox.SelectedItems.Cast<string>().ToList();
            if (items.Contains("All") || items.Contains("All Weeks (Total)")) return "All";
            return string.Join(", ", items);
        }

        private void ExportWeeks_Click(object sender, RoutedEventArgs e)
        {
            if (gridHistory.ItemsSource is not List<WeeklyPivotRow> data || data.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export Info");
                return;
            }

            string skillHeader = $"Skills: {GetSelectedItemsString(lstTeams)}";

            SaveToExcel(workbook =>
            {
                var ws = workbook.Worksheets.Add("Weekly Volume");

                ws.Cell(1, 1).Value = skillHeader;
                ws.Range(1, 1, 1, 9).Merge().Style.Font.Bold = true;

                int startRow = 3;
                string[] headers = { "Week", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Total" };

                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cell(startRow, i + 1).Value = headers[i];
                }

                var headerRange = ws.Range(startRow, 1, startRow, 9);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                for (int i = 0; i < data.Count; i++)
                {
                    var row = data[i];
                    int r = startRow + 1 + i;

                    ws.Cell(r, 1).Value = row.WeekLabel;
                    ws.Cell(r, 2).Value = row.Sun;
                    ws.Cell(r, 3).Value = row.Mon;
                    ws.Cell(r, 4).Value = row.Tue;
                    ws.Cell(r, 5).Value = row.Wed;
                    ws.Cell(r, 6).Value = row.Thu;
                    ws.Cell(r, 7).Value = row.Fri;
                    ws.Cell(r, 8).Value = row.Sat;
                    ws.Cell(r, 9).Value = row.WeekTotal;
                }

                ws.Columns().AdjustToContents();

            }, "Weekly_Volume_Export");
        }

        private void ExportIntervals_Click(object sender, RoutedEventArgs e)
        {
            if (gridIntervals.ItemsSource is not List<ArrivalIntervalRow> data || data.Count == 0)
            {
                MessageBox.Show("No data to export.", "Export Info");
                return;
            }

            string skillHeader = $"Skills: {GetSelectedItemsString(lstTeams)}";
            string weekHeader = $"Weeks: {GetSelectedItemsString(lstWeeks)}";
            bool isPercent = rbPattern.IsChecked == true;

            double totalSun = 0, totalMon = 0, totalTue = 0, totalWed = 0, totalThu = 0, totalFri = 0, totalSat = 0;
            if (isPercent)
            {
                foreach (var r in data.Where(x => x.Label != "TOTAL"))
                {
                    totalSun += r.SunVal; totalMon += r.MonVal; totalTue += r.TueVal;
                    totalWed += r.WedVal; totalThu += r.ThuVal; totalFri += r.FriVal; totalSat += r.SatVal;
                }
            }

            SaveToExcel(workbook =>
            {
                var ws = workbook.Worksheets.Add("Interval Distribution");

                int totalCols = isPercent ? 8 : 9;

                ws.Cell(1, 1).Value = skillHeader;
                ws.Range(1, 1, 1, totalCols).Merge().Style.Font.Bold = true;

                ws.Cell(2, 1).Value = weekHeader;
                ws.Range(2, 1, 2, totalCols).Merge().Style.Font.Bold = true;

                int startRow = 4;
                var headerList = new List<string> { "Interval", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };

                if (!isPercent) headerList.Add("Total");

                for (int i = 0; i < headerList.Count; i++)
                {
                    ws.Cell(startRow, i + 1).Value = headerList[i];
                }

                var headerRange = ws.Range(startRow, 1, startRow, headerList.Count);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                for (int i = 0; i < data.Count; i++)
                {
                    var row = data[i];
                    int r = startRow + 1 + i;

                    ws.Cell(r, 1).Value = row.Label;

                    void WriteCell(int col, double val, double dayTotal)
                    {
                        var cell = ws.Cell(r, col);
                        if (isPercent)
                        {
                            double pct = dayTotal > 0 ? (val / dayTotal) : 0;
                            cell.Value = pct;
                            cell.Style.NumberFormat.Format = "0.0%";
                        }
                        else
                        {
                            cell.Value = val;
                            cell.Style.NumberFormat.Format = "#,##0";
                        }
                    }

                    WriteCell(2, row.SunVal, totalSun);
                    WriteCell(3, row.MonVal, totalMon);
                    WriteCell(4, row.TueVal, totalTue);
                    WriteCell(5, row.WedVal, totalWed);
                    WriteCell(6, row.ThuVal, totalThu);
                    WriteCell(7, row.FriVal, totalFri);
                    WriteCell(8, row.SatVal, totalSat);

                    if (!isPercent)
                    {
                        double rowSum = row.SunVal + row.MonVal + row.TueVal + row.WedVal +
                                        row.ThuVal + row.FriVal + row.SatVal;

                        var cellTotal = ws.Cell(r, 9);
                        cellTotal.Value = rowSum;
                        cellTotal.Style.NumberFormat.Format = "#,##0";
                        cellTotal.Style.Font.Bold = true;
                    }

                    if (row.Label == "TOTAL")
                    {
                        ws.Row(r).Style.Font.Bold = true;
                        ws.Row(r).Style.Fill.BackgroundColor = XLColor.WhiteSmoke;
                    }
                }

                ws.Columns().AdjustToContents();

            }, "Interval_Distribution_Export");
        }

        private void SaveToExcel(Action<XLWorkbook> buildWorkbookAction, string defaultName)
        {
            SaveFileDialog dlg = new SaveFileDialog
            {
                FileName = $"{defaultName}_{DateTime.Now:yyyyMMdd}",
                DefaultExt = ".xlsx",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    using (var workbook = new XLWorkbook())
                    {
                        buildWorkbookAction(workbook);
                        workbook.SaveAs(dlg.FileName);
                    }
                    MessageBox.Show("Export Successful!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving file: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}