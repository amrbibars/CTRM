using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CTRM.DB;

namespace CTRM.Services
{
    public partial class Forecasting : Window
    {
        private List<string> _allSkills = new List<string>();

        // Variables to hold the data for the Export function
        private List<string> _lastSelectedSkills;
        private DataTable _dtVolumeOutput;
        private List<ArrivalPatternData> _patternList;
        private DataTable _dtIntervalOutput;

        public Forecasting()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
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
                    var loader = new IBV_Historical_Loader();
                    var loadedData = loader.Load(openFileDialog.FileName);

                    if (loadedData != null && loadedData.Count > 0)
                    {
                        DataManager.Instance.SetData(loadedData);
                        MessageBox.Show("Data successfully loaded.", "Load Success", MessageBoxButton.OK, MessageBoxImage.Information);
                        UpdateDataStatus();
                        btnExport.IsEnabled = false; // Reset export button
                    }
                    else
                    {
                        MessageBox.Show("No valid data found.", "Load Error", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                var data = DataManager.Instance.AllData;
                txtDataStatus.Text = $"Status: Data Loaded ({data.Count:N0} records active)";
                txtDataStatus.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#27AE60"));

                _allSkills = data.Select(x => x.Team).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList();
                lstSkills.ItemsSource = _allSkills;

                int totalWeeksInData = data.Select(x => x.WeekNumber).Distinct().Count();
                int maxWeeks = Math.Min(12, totalWeeksInData);
                int minWeeks = Math.Min(3, maxWeeks);

                cmbLookupRange.Items.Clear();
                if (maxWeeks > 0)
                {
                    for (int i = minWeeks; i <= maxWeeks; i++) cmbLookupRange.Items.Add(i);
                    cmbLookupRange.SelectedIndex = cmbLookupRange.Items.Count - 1;
                }
            }
            else
            {
                txtDataStatus.Text = "Status: No data loaded. Please upload a file.";
                txtDataStatus.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E74C3C"));
                lstSkills.ItemsSource = null;
                cmbLookupRange.Items.Clear();
            }
        }

        private void txtSkillSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (lstSkills == null) return;

            if (string.IsNullOrWhiteSpace(txtSkillSearch.Text))
                lstSkills.ItemsSource = _allSkills;
            else
                lstSkills.ItemsSource = _allSkills.Where(s => s.ToLower().Contains(txtSkillSearch.Text.ToLower())).ToList();
        }

        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            if (!DataManager.Instance.HasData) return;

            var selectedSkills = lstSkills.SelectedItems.Cast<string>().ToList();
            if (selectedSkills.Count == 0)
            {
                MessageBox.Show("Please select at least one skill.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int lookupRange = cmbLookupRange.SelectedItem != null ? (int)cmbLookupRange.SelectedItem : 0;
            if (lookupRange == 0) return;

            double recencyLevel = sldRecency.Value;
            bool ignoreTrend = chkIgnoreTrend.IsChecked == true;

            List<int> excludedWeeks = new List<int>();
            if (!string.IsNullOrWhiteSpace(txtExclusions.Text))
            {
                var parts = txtExclusions.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                    if (int.TryParse(part.Trim(), out int w)) excludedWeeks.Add(w);
            }

            try
            {
                _lastSelectedSkills = selectedSkills;
                RunForecastEngine(selectedSkills, excludedWeeks, lookupRange, recencyLevel, ignoreTrend);

                // Enable Export once calculations succeed
                btnExport.IsEnabled = true;
                MessageBox.Show("Forecast and Interval Distribution calculated successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during forecast calculation: {ex.Message}", "Calculation Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RunForecastEngine(List<string> skills, List<int> excludedWeeks, int range, double recencyLevel, bool ignoreTrend)
        {
            var baseData = DataManager.Instance.AllData
                .Where(x => skills.Contains(x.Team) && !excludedWeeks.Contains(x.WeekNumber)).ToList();

            if (baseData.Count == 0) throw new Exception("No data available after applying filters.");

            var targetWeeks = baseData.Select(x => x.WeekNumber).Distinct()
                .OrderByDescending(w => w).Take(range).OrderBy(w => w).ToList();

                var dailyData = baseData.Where(x => targetWeeks.Contains(x.WeekNumber))
                .GroupBy(x => new {
                    DateOnly = x.Date.Date,
                    DayOfWeek = x.Date.DayOfWeek,
                    WeekNum = x.WeekNumber,
                    Start = x.StartTime
                })
                .Select(g => new {
                    Date = g.Key.DateOnly,
                    DOW = g.Key.DayOfWeek,
                    WeekNum = g.Key.WeekNum,
                    Interval = g.Key.Start.ToString(@"hh\:mm"),
                    Volume = g.Sum(c => (double)c.Offered)
                }).ToList();

            var weeklyVolumes = dailyData.GroupBy(x => x.WeekNum).ToDictionary(g => g.Key, g => g.Sum(x => x.Volume));
            var dailyVolumes = dailyData.GroupBy(x => x.Date).ToDictionary(g => g.Key, g => g.Sum(x => x.Volume));

            // --- STAGE 1: DOW PATTERN ---
            var dailyPercentages = dailyVolumes.Select(d => new {
                Date = d.Key,
                DOW = d.Key.DayOfWeek,
                WeekNum = dailyData.First(x => x.Date == d.Key).WeekNum,
                Pct = weeklyVolumes[dailyData.First(x => x.Date == d.Key).WeekNum] > 0 ? (d.Value / weeklyVolumes[dailyData.First(x => x.Date == d.Key).WeekNum]) : 0
            }).ToList();

            DayOfWeek[] dowOrder = { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday };
            _patternList = new List<ArrivalPatternData>();
            double totalNormalAvg = 0;

            foreach (var day in dowOrder)
            {
                var dayPcts = dailyPercentages.Where(x => x.DOW == day).Select(x => x.Pct).ToList();
                double avg = dayPcts.Any() ? dayPcts.Average() : 0;
                double stdDev = CalculateStandardDeviation(dayPcts);
                double ll = avg - stdDev; double ul = avg + stdDev;

                var normalPcts = dayPcts.Where(p => p >= ll && p <= ul).ToList();
                double normalAvg = normalPcts.Any() ? normalPcts.Average() : avg;
                totalNormalAvg += normalAvg;

                _patternList.Add(new ArrivalPatternData { DOW = day.ToString(), Average = avg, StdDev = stdDev, LowerLimit = ll, UpperLimit = ul, NormalAverage = normalAvg });
            }

            foreach (var p in _patternList) p.NormalAverage = totalNormalAvg > 0 ? (p.NormalAverage / totalNormalAvg) : (1.0 / 7.0);

            double[] weights = GetDynamicWeights(targetWeeks.Count, recencyLevel);
            double weightedForecast = 0;
            for (int i = 0; i < targetWeeks.Count; i++) weightedForecast += weeklyVolumes[targetWeeks[i]] * weights[i];

            double trendForecast = 0;
            if (!ignoreTrend && targetWeeks.Count > 1)
            {
                List<double> ratios = new List<double>();
                for (int i = 1; i < targetWeeks.Count; i++)
                {
                    double prev = weeklyVolumes[targetWeeks[i - 1]];
                    double curr = weeklyVolumes[targetWeeks[i]];
                    if (prev > 0) ratios.Add(curr / prev);
                }
                double avgTrend = ratios.Any() ? ratios.Average() : 1.0;
                trendForecast = weeklyVolumes[targetWeeks.Last()] * avgTrend;
            }

            double finalWeeklyVolume = ignoreTrend || targetWeeks.Count <= 1 ? weightedForecast : (weightedForecast + trendForecast) / 2.0;

            // --- STAGE 2: INTERVAL NORMALIZATION ---
            var intervalPcts = dailyData.Select(x => new {
                x.Date,
                x.DOW,
                x.Interval,
                Pct = dailyVolumes[x.Date] > 0 ? x.Volume / dailyVolumes[x.Date] : 0
            }).ToList();

            var intervalStats = intervalPcts.GroupBy(x => new { x.DOW, x.Interval })
                .Select(g => {
                    double iAvg = g.Average(x => x.Pct);
                    double iStd = CalculateStandardDeviation(g.Select(x => x.Pct).ToList());
                    double iLl = iAvg - iStd; double iUl = iAvg + iStd;
                    var normalValues = g.Where(x => x.Pct >= iLl && x.Pct <= iUl).Select(x => x.Pct).ToList();
                    return new { g.Key.DOW, g.Key.Interval, NormalAvg = normalValues.Any() ? normalValues.Average() : iAvg };
                }).ToList();

            var dowIntervalSums = intervalStats.GroupBy(x => x.DOW).ToDictionary(g => g.Key, g => g.Sum(x => x.NormalAvg));

            // Generate Top Grid (Daily Volume)
            _dtVolumeOutput = new DataTable();
            _dtVolumeOutput.Columns.Add("Period", typeof(string));
            foreach (var day in dowOrder) _dtVolumeOutput.Columns.Add(day.ToString(), typeof(int));
            _dtVolumeOutput.Columns.Add("Total", typeof(int));

            foreach (var week in targetWeeks)
            {
                DataRow row = _dtVolumeOutput.NewRow();
                row["Period"] = $"Week {week} Actual";
                foreach (var day in dowOrder) row[day.ToString()] = (int)Math.Round(dailyVolumes.Where(x => x.Key.DayOfWeek == day && dailyData.Any(d => d.Date == x.Key && d.WeekNum == week)).Sum(x => x.Value));
                row["Total"] = (int)Math.Round(weeklyVolumes[week]);
                _dtVolumeOutput.Rows.Add(row);
            }

            DataRow forecastRow = _dtVolumeOutput.NewRow();
            forecastRow["Period"] = "Forecasted Volume";
            int forecastTotal = 0;
            Dictionary<string, int> forecastedDailyVols = new Dictionary<string, int>();

            foreach (var day in dowOrder)
            {
                int dayForecast = (int)Math.Round(finalWeeklyVolume * _patternList.First(p => p.DOW == day.ToString()).NormalAverage);
                forecastRow[day.ToString()] = dayForecast;
                forecastedDailyVols[day.ToString()] = dayForecast;
                forecastTotal += dayForecast;
            }
            forecastRow["Total"] = forecastTotal;
            _dtVolumeOutput.Rows.Add(forecastRow);

            // Generate Interval Grid (Hidden in UI, Ready for Export)
            _dtIntervalOutput = new DataTable();
            _dtIntervalOutput.Columns.Add("Interval", typeof(string));
            foreach (var day in dowOrder) _dtIntervalOutput.Columns.Add(day.ToString(), typeof(int));
            _dtIntervalOutput.Columns.Add("Total", typeof(int));

            List<string> timeSlots = GenerateTimeSlots();
            foreach (var time in timeSlots)
            {
                DataRow row = _dtIntervalOutput.NewRow();
                row["Interval"] = time;
                int rowTotal = 0;

                foreach (var day in dowOrder)
                {
                    var stat = intervalStats.FirstOrDefault(x => x.DOW == day && x.Interval == time);
                    double finalPct = stat != null && dowIntervalSums[day] > 0 ? (stat.NormalAvg / dowIntervalSums[day]) : 0;

                    int intVol = (int)Math.Round(forecastedDailyVols[day.ToString()] * finalPct);
                    row[day.ToString()] = intVol;
                    rowTotal += intVol;
                }
                row["Total"] = rowTotal;
                _dtIntervalOutput.Rows.Add(row);
            }

            // Bind to UI
            dgVolumeOutput.ItemsSource = _dtVolumeOutput.DefaultView;
            dgPatternOutput.ItemsSource = _patternList;
            dgIntervalOutput.ItemsSource = _dtIntervalOutput.DefaultView;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_dtVolumeOutput == null || _dtIntervalOutput == null || _patternList == null) return;

            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel CSV (*.csv)|*.csv",
                FileName = $"Forecast_Export_{DateTime.Now:yyyyMMdd_HHmm}.csv",
                Title = "Export Forecast to Excel"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    StringBuilder sb = new StringBuilder();

                    // 1. SKILLS
                    sb.AppendLine("--- SELECTED SKILLS ---");
                    sb.AppendLine(string.Join(" | ", _lastSelectedSkills));
                    sb.AppendLine();

                    // 2. DAILY & WEEKLY VOLUME
                    sb.AppendLine("--- DAILY & WEEKLY VOLUME FORECAST ---");
                    sb.AppendLine(string.Join(",", _dtVolumeOutput.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                    foreach (DataRow row in _dtVolumeOutput.Rows)
                        sb.AppendLine(string.Join(",", row.ItemArray));
                    sb.AppendLine();

                    // 3. DOW ARRIVAL PATTERN
                    sb.AppendLine("--- NORMALIZED ARRIVAL PATTERN ---");
                    sb.AppendLine("DOW,Average %,Std Dev %,Lower Limit %,Upper Limit %,Normal Avg %");
                    foreach (var p in _patternList)
                        sb.AppendLine($"{p.DOW},{p.Average:P2},{p.StdDev:P2},{p.LowerLimit:P2},{p.UpperLimit:P2},{p.NormalAverage:P2}");
                    sb.AppendLine();

                    // 4. INTERVALS
                    sb.AppendLine("--- INTERVAL FORECAST ---");
                    sb.AppendLine(string.Join(",", _dtIntervalOutput.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));
                    foreach (DataRow row in _dtIntervalOutput.Rows)
                        sb.AppendLine(string.Join(",", row.ItemArray));

                    File.WriteAllText(saveDialog.FileName, sb.ToString());
                    MessageBox.Show("Export complete! You can open this file directly in Excel.", "Export Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error saving file: " + ex.Message, "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private List<string> GenerateTimeSlots()
        {
            var slots = new List<string>();
            for (int h = 0; h < 24; h++)
            {
                slots.Add($"{h:D2}:00");
                slots.Add($"{h:D2}:30");
            }
            return slots;
        }

        private double[] GetDynamicWeights(int count, double sliderValue)
        {
            if (count == 0) return new double[0];
            if (count == 1) return new double[] { 1.0 };

            double[] weights = new double[count];
            double bias = 0.2;
            double effectiveBias = bias + (sliderValue / 10.0);

            for (int i = 0; i < count; i++)
            {
                double pos = (2.0 * i / (count - 1)) - 1.0;
                weights[i] = Math.Exp(effectiveBias * 3.0 * pos);
            }

            double sum = weights.Sum();
            for (int i = 0; i < count; i++) weights[i] /= sum;

            return weights;
        }

        private double CalculateStandardDeviation(List<double> values)
        {
            if (values.Count <= 1) return 0;
            double avg = values.Average();
            return Math.Sqrt(values.Sum(val => Math.Pow(val - avg, 2)) / (values.Count - 1));
        }

        public class ArrivalPatternData
        {
            public string DOW { get; set; }
            public double Average { get; set; }
            public double StdDev { get; set; }
            public double LowerLimit { get; set; }
            public double UpperLimit { get; set; }
            public double NormalAverage { get; set; }
        }
    }
}