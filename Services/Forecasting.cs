using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using CTRM.DB;

namespace CTRM.Services
{
    public partial class Forecasting : Window
    {
        private List<string> _allSkills = new List<string>();

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

                // Populate Skills List 
                _allSkills = data.Select(x => x.Team).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().OrderBy(s => s).ToList();
                lstSkills.ItemsSource = _allSkills;

                // Dynamically Populate Lookup Range
                int totalWeeksInData = data.Select(x => x.WeekNumber).Distinct().Count();
                int maxWeeks = Math.Min(12, totalWeeksInData);
                int minWeeks = Math.Min(3, maxWeeks);

                cmbLookupRange.Items.Clear();
                if (maxWeeks > 0)
                {
                    for (int i = minWeeks; i <= maxWeeks; i++)
                    {
                        cmbLookupRange.Items.Add(i);
                    }
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
            {
                lstSkills.ItemsSource = _allSkills;
            }
            else
            {
                var filter = txtSkillSearch.Text.ToLower();
                lstSkills.ItemsSource = _allSkills.Where(s => s.ToLower().Contains(filter)).ToList();
            }
        }

        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            if (!DataManager.Instance.HasData)
            {
                MessageBox.Show("Please load data before forecasting.", "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

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

            // Parse Exclusions
            List<int> excludedWeeks = new List<int>();
            if (!string.IsNullOrWhiteSpace(txtExclusions.Text))
            {
                var parts = txtExclusions.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    if (int.TryParse(part.Trim(), out int w)) excludedWeeks.Add(w);
                }
            }

            try
            {
                RunForecastEngine(selectedSkills, excludedWeeks, lookupRange, recencyLevel, ignoreTrend);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during forecast calculation: {ex.Message}", "Calculation Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// THE CORE MATH ENGINE
        /// </summary>
        private void RunForecastEngine(List<string> skills, List<int> excludedWeeks, int range, double recencyLevel, bool ignoreTrend)
        {
            // 1. Filter Data
            var baseData = DataManager.Instance.AllData
                .Where(x => skills.Contains(x.Team))
                .Where(x => !excludedWeeks.Contains(x.WeekNumber))
                .ToList();

            if (baseData.Count == 0) throw new Exception("No data available after applying skill and week exclusion filters.");

            // 2. Identify Target Weeks (Last N weeks chronologically)
            var targetWeeks = baseData.Select(x => x.WeekNumber).Distinct()
                .OrderByDescending(w => w).Take(range).OrderBy(w => w).ToList();

            // 3. Aggregate Daily and Weekly Volumes
            var dailyData = baseData.Where(x => targetWeeks.Contains(x.WeekNumber))
                .GroupBy(x => x.Date.Date)
                .Select(g => new {
                    Date = g.Key,
                    DOW = g.Key.DayOfWeek,
                    WeekNum = g.First().WeekNumber, // Assuming chronological data aligns perfectly
                    Volume = g.Sum(x => (double)x.Offered)
                }).ToList();

            var weeklyVolumes = dailyData.GroupBy(x => x.WeekNum)
                .ToDictionary(g => g.Key, g => g.Sum(x => x.Volume));

            // Calculate Daily Percentages
            var dailyPercentages = dailyData.Select(d => new {
                d.Date,
                d.DOW,
                d.WeekNum,
                Pct = weeklyVolumes[d.WeekNum] > 0 ? (d.Volume / weeklyVolumes[d.WeekNum]) : 0
            }).ToList();

            // 4. Calculate Normalized Arrival Pattern (Step 1B)
            DayOfWeek[] daysOfWeek = { DayOfWeek.Sunday, DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday, DayOfWeek.Saturday };
            var patternList = new List<ArrivalPatternData>();
            double totalNormalAvg = 0;

            foreach (var day in daysOfWeek)
            {
                var dayPcts = dailyPercentages.Where(x => x.DOW == day).Select(x => x.Pct).ToList();
                double avg = dayPcts.Any() ? dayPcts.Average() : 0;
                double stdDev = CalculateStandardDeviation(dayPcts);
                double ll = avg - stdDev;
                double ul = avg + stdDev;

                // Filter to Normalized Range
                var normalPcts = dayPcts.Where(p => p >= ll && p <= ul).ToList();
                double normalAvg = normalPcts.Any() ? normalPcts.Average() : avg; // fallback to avg if none fall in range

                totalNormalAvg += normalAvg;

                patternList.Add(new ArrivalPatternData
                {
                    DOW = day.ToString(),
                    Average = avg,
                    StdDev = stdDev,
                    LowerLimit = ll,
                    UpperLimit = ul,
                    NormalAverage = normalAvg
                });
            }

            // Rescale Normal Averages to exactly 100%
            foreach (var p in patternList)
            {
                p.NormalAverage = totalNormalAvg > 0 ? (p.NormalAverage / totalNormalAvg) : (1.0 / 7.0);
            }

            // 5. Calculate Weighted Average Forecast
            double[] weights = GetDynamicWeights(targetWeeks.Count, recencyLevel);
            double weightedForecast = 0;
            for (int i = 0; i < targetWeeks.Count; i++)
            {
                weightedForecast += weeklyVolumes[targetWeeks[i]] * weights[i];
            }

            // 6. Calculate Trend Forecast
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

            // 7. Final Volume Calculation
            double finalWeeklyVolume = ignoreTrend || targetWeeks.Count <= 1
                ? weightedForecast
                : (weightedForecast + trendForecast) / 2.0;

            // 8. Bind to Grids
            RenderOutputGrids(targetWeeks, weeklyVolumes, dailyData, patternList, finalWeeklyVolume, daysOfWeek);
        }

        private void RenderOutputGrids(List<int> weeks, Dictionary<int, double> weeklyVols, dynamic dailyData, List<ArrivalPatternData> patternList, double finalWeeklyVol, DayOfWeek[] dowOrder)
        {
            DataTable dtVolume = new DataTable();
            dtVolume.Columns.Add("Period", typeof(string));

            foreach (var day in dowOrder) dtVolume.Columns.Add(day.ToString(), typeof(int));
            dtVolume.Columns.Add("Total", typeof(int));

            // Add Actuals Rows
            foreach (var week in weeks)
            {
                DataRow row = dtVolume.NewRow();
                row["Period"] = $"Week {week} Actual";

                foreach (var day in dowOrder)
                {
                    // Find the volume for this specific day in this specific week
                    var dayVol = ((IEnumerable<dynamic>)dailyData)
                                 .FirstOrDefault(x => x.WeekNum == week && x.DOW == day)?.Volume ?? 0;
                    row[day.ToString()] = (int)Math.Round(dayVol);
                }
                row["Total"] = (int)Math.Round(weeklyVols[week]);
                dtVolume.Rows.Add(row);
            }

            // Add Forecast Row
            DataRow forecastRow = dtVolume.NewRow();
            forecastRow["Period"] = "Forecasted Volume";
            int forecastTotal = 0;

            foreach (var day in dowOrder)
            {
                double dayPct = patternList.First(p => p.DOW == day.ToString()).NormalAverage;
                int dayForecast = (int)Math.Round(finalWeeklyVol * dayPct);
                forecastRow[day.ToString()] = dayForecast;
                forecastTotal += dayForecast;
            }
            forecastRow["Total"] = forecastTotal;
            dtVolume.Rows.Add(forecastRow);

            dgVolumeOutput.ItemsSource = dtVolume.DefaultView;
            dgPatternOutput.ItemsSource = patternList;
        }

        /// <summary>
        /// Exponential weight generator based on Slider (-5 to +5).
        /// Mathematically smooths the weights so they always equal 100%.
        /// </summary>
        private double[] GetDynamicWeights(int count, double sliderValue)
        {
            if (count == 0) return new double[0];
            if (count == 1) return new double[] { 1.0 };

            double[] weights = new double[count];

            // Bias maps the -5 to +5 slider to a curve exponent (-0.3 to 0.7)
            double bias = 0.2; // Natural bias (mimics 40/40/20)
            double effectiveBias = bias + (sliderValue / 10.0);

            for (int i = 0; i < count; i++)
            {
                // Normalize index: -1 (Oldest Week) to +1 (Newest Week)
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
            double sumOfSquares = values.Sum(val => Math.Pow(val - avg, 2));
            return Math.Sqrt(sumOfSquares / (values.Count - 1));
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