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

                // Sanitize header for Excel Filters and Quotes
                headerLine = headerLine.Replace("\uFEFF", "").Replace("\"", "").Trim();
                var headers = headerLine.Split(',').Select(h => h.Trim().ToUpper()).ToList();

                int idxDate = headers.IndexOf("LABEL_YYYY_MM_DD");
                int idxInterval = headers.IndexOf("LABEL_YYYY_MM_DD_HH24_30INT");
                int idxOffered = headers.IndexOf("OFFERED");
                int idxTeam = headers.IndexOf("RESOURCE_NAME");
                int idxMedia = headers.IndexOf("MEDIA_NAME");
                int idxType = headers.IndexOf("INTERACTION_TYPE");
                int idxAht = headers.IndexOf("AHT");

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

                        // Use double parse for safety with Excel formats
                        int offered = (int)double.Parse(columns[idxOffered]);
                        int aht = (idxAht != -1 && idxAht < columns.Length) ? (int)double.Parse(columns[idxAht]) : 0;
                        string team = (idxTeam != -1 && idxTeam < columns.Length) ? columns[idxTeam].Trim() : "Unknown";

                        var weekInfo = GetFiscalWeek(date);

                        results.Add(new IBV_CallData
                        {
                            Date = date,
                            StartTime = startTime,
                            Team = team,
                            Offered = offered,
                            Aht = aht,
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
        private List<IBV_CallData> _allData;
        private List<IBV_CallData> _filteredData;
        private bool _isDataLoading = false;

        public Arrival_Pattern()
        {
            InitializeComponent();
            _allData = new List<IBV_CallData>();
            _filteredData = new List<IBV_CallData>();
        }

        private void LoadData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "IB Voice Reports|*.csv;*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var loader = new IBV_Historical_Loader();
                    _allData = loader.Load(openFileDialog.FileName);

                    if (_allData.Count > 0)
                    {
                        _isDataLoading = true;

                        // Populate Team Dropdown with "All" option
                        var teams = _allData.Select(x => x.Team).Distinct().OrderBy(t => t).ToList();
                        teams.Insert(0, "All");
                        cmbTeams.ItemsSource = teams;

                        if (teams.Count > 0) cmbTeams.SelectedIndex = 0;

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
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void OnFilterChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isDataLoading) return;
            RefreshDashboard();
        }

        private void RefreshDashboard()
        {
            if (_allData == null || _allData.Count == 0 || cmbTeams.SelectedItem == null) return;

            string? selectedTeam = cmbTeams.SelectedItem.ToString();

            if (selectedTeam == "All")
                _filteredData = _allData;
            else
                _filteredData = _allData.Where(x => x.Team == selectedTeam).ToList();

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

            string viewMode = (cmbView.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "All Days (Trend)";
            List<double> vals = new List<double>();

            if (viewMode == "All Days (Trend)")
            {
                foreach (var r in data) { vals.AddRange(new double[] { r.Mon, r.Tue, r.Wed, r.Thu, r.Fri, r.Sat, r.Sun }); }
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
                        DayOfWeek.Monday => (double)r.Sun,
                        DayOfWeek.Tuesday => (double)r.Mon,
                        DayOfWeek.Wednesday => (double)r.Tue,
                        DayOfWeek.Thursday => (double)r.Wed,
                        DayOfWeek.Friday => (double)r.Thu,
                        DayOfWeek.Saturday => (double)r.Fri,
                        _ => (double)r.Sat
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
    }
}