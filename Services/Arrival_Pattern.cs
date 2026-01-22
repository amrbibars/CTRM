using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32; // For OpenFileDialog

namespace CTRM
{
    // ---------------------------------------------------------
    // 1. DATA STRUCTURE (Lightweight)
    // ---------------------------------------------------------
    public struct IBV_CallData
    {
        public DateTime Date;
        public TimeSpan StartTime; // e.g., 06:30
        public string Team;        // RESOURCE_NAME
        public int Offered;        // Calls Volume
        public int Aht;            // Handle Time
        public int WeekNumber;     // Custom Fiscal Week
        public int Year;           // Custom Fiscal Year
    }

    // ---------------------------------------------------------
    // 2. LOGIC ENGINE (Specific to IBV CSV/Excel Format)
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

                // Normalize headers to Upper Case for safe matching
                var headers = headerLine.Split(',').Select(h => h.Trim().ToUpper()).ToList();

                // Dynamic Column Mapping
                int idxDate = headers.IndexOf("LABEL_YYYY_MM_DD");
                int idxInterval = headers.IndexOf("LABEL_YYYY_MM_DD_HH24_30INT");
                int idxOffered = headers.IndexOf("OFFERED");
                int idxTeam = headers.IndexOf("RESOURCE_NAME");
                int idxMedia = headers.IndexOf("MEDIA_NAME");
                int idxType = headers.IndexOf("INTERACTION_TYPE");
                int idxAht = headers.IndexOf("AHT");

                // Basic Validation
                if (idxDate == -1 || idxInterval == -1 || idxOffered == -1)
                    throw new Exception("Missing standard columns (Date, Interval, or Offered).");

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var columns = line.Split(',');

                    // Strict Filtering: Voice & Inbound Only
                    if (idxMedia != -1 && columns[idxMedia].Trim() != "Voice") continue;
                    if (idxType != -1 && columns[idxType].Trim() != "Inbound") continue;

                    try
                    {
                        // 1. Parse Date
                        if (!DateTime.TryParseExact(columns[idxDate],
                            new[] { "M/d/yyyy", "MM/dd/yyyy", "M/dd/yyyy", "MM/d/yyyy" },
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                            continue;

                        // 2. Parse Interval (Get "06:30" from "2025-12-15 06:30-07:00")
                        string timePart = columns[idxInterval].Split(' ')[1];
                        TimeSpan startTime = TimeSpan.Parse(timePart.Split('-')[0]);

                        // 3. Parse Numbers
                        int offered = int.Parse(columns[idxOffered]);
                        int aht = (idxAht != -1) ? int.Parse(columns[idxAht]) : 0;
                        string team = columns[idxTeam];

                        // 4. Calculate Fiscal Week (Sunday Start)
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
                    catch
                    {
                        // Skip bad rows silently
                        continue;
                    }
                }
            }
            return results;
        }

        // Logic: Sunday Start. If week has Jan days, it's Week 1. No Week 53.
        private (int Year, int Week) GetFiscalWeek(DateTime date)
        {
            DateTime startOfWeek = date.AddDays(-(int)date.DayOfWeek); // Sunday
            DateTime endOfWeek = startOfWeek.AddDays(6);               // Saturday

            int fiscalYear = startOfWeek.Year;

            // "January Rule": If the week bridges years, it belongs to the new year (Week 1)
            if (endOfWeek.Year > startOfWeek.Year)
            {
                return (endOfWeek.Year, 1);
            }

            // Calculate Week Number from the first Sunday of the year
            DateTime jan1 = new DateTime(fiscalYear, 1, 1);
            // Adjust Jan 1 to find the Sunday that started its week
            DateTime week1Start = jan1.AddDays(-(int)jan1.DayOfWeek);

            int daysDiff = (startOfWeek - week1Start).Days;
            int weekNum = (daysDiff / 7) + 1;

            return (fiscalYear, weekNum);
        }
    }


    // ---------------------------------------------------------
    // 3. ANALYZER ENGINE (Calculates the Pattern)
    // ---------------------------------------------------------
    public class IBV_Analyzer
    {
        // Structure for the final "Average" result
        public struct IntervalResult
        {
            public DayOfWeek Day;
            public TimeSpan Time;
            public double AvgOffered;
            public double AvgAHT;
        }

        public List<IntervalResult> CalculateAveragePattern(List<IBV_CallData> rawData)
        {
            var results = new List<IntervalResult>();

            // 1. Group by Day of Week and Time Interval
            //    Key = "Monday-08:30", Value = List of rows
            var groupedData = rawData
                .GroupBy(x => new { x.Date.DayOfWeek, x.StartTime })
                .OrderBy(g => g.Key.DayOfWeek)
                .ThenBy(g => g.Key.StartTime);

            foreach (var group in groupedData)
            {
                // 2. Count distinct weeks for this specific interval
                //    (Crucial: If one Monday was a holiday and missing, don't divide by total weeks)
                int weeksCount = group.Select(x => x.WeekNumber).Distinct().Count();

                if (weeksCount == 0) continue;

                // 3. Calculate Averages
                double totalOffered = group.Sum(x => x.Offered);

                // Weighted AHT Formula: Sum(Offered * AHT) / TotalOffered
                // This prevents short calls from skewing the average
                double totalWorkload = group.Sum(x => x.Offered * x.Aht);
                double weightedAHT = (totalOffered > 0) ? (totalWorkload / totalOffered) : 0;

                results.Add(new IntervalResult
                {
                    Day = group.Key.DayOfWeek,
                    Time = group.Key.StartTime,
                    AvgOffered = totalOffered / weeksCount, // Simple Average
                    AvgAHT = weightedAHT
                });
            }

            return results;
        }
    }

    // ---------------------------------------------------------
    // 4. MAIN WINDOW LOGIC
    // ---------------------------------------------------------
    public partial class Arrival_Pattern : Window
    {
        // Store loaded data in memory for calculations
        private List<IBV_CallData> _loadedData;

        public Arrival_Pattern()
        {
            InitializeComponent();
            _loadedData = new List<IBV_CallData>();
        }

        private void LoadData_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "IB Voice Reports|*.csv;*.xlsx;*.xls";
            openFileDialog.Title = "Select Inbound Voice Data (4-12 Weeks)";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    // Call the internal class
                    var loader = new IBV_Historical_Loader();
                    _loadedData = loader.Load(openFileDialog.FileName);

                    if (_loadedData.Count > 0)
                    {
                        MessageBox.Show($"Data Loaded Successfully!\nRows: {_loadedData.Count}\n" +
                                        $"Date Range: {_loadedData.Min(x => x.Date):d} to {_loadedData.Max(x => x.Date):d}",
                                        "Import Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("No 'Voice' / 'Inbound' data found in file.", "Empty Load", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error processing file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}