using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace CTRM.Services
{
    public class BudgetingCalculator
    {
        public DataTable CalculateForecast(DataTable historicalData,
                                           int startMonth,
                                           int startYear,
                                           int endMonth,
                                           int endYear)
        {
            // Clone structure for the forecast table
            DataTable forecastTable = historicalData.Clone();

            // We need to identify Market columns (indices 2 to N)
            // Assuming Column 0 is Year, Column 1 is Month
            List<int> marketColIndices = [];
            for (int i = 2; i < historicalData.Columns.Count; i++)
            {
                if (historicalData.Columns[i].DataType == typeof(double) ||
                    historicalData.Columns[i].DataType == typeof(int) ||
                    historicalData.Columns[i].DataType == typeof(decimal))
                {
                    marketColIndices.Add(i);
                }
            }

            // Convert DataTable to a more usable Dictionary structure: [Year, Month] -> RowData
            var historyMap = new Dictionary<(int Year, int Month), DataRow>();
            foreach (DataRow row in historicalData.Rows)
            {
                int y = Convert.ToInt32(row["Year"]);
                int m = DateTime.ParseExact(row["Month"]?.ToString()?? "", "MMM", System.Globalization.CultureInfo.InvariantCulture).Month;
                historyMap[(y, m)] = row;
            }

            // --- Simulation Loop ---
            // We start from the user's defined Start Month/Year and roll forward until End Month/Year

            DateTime currentDate = new(startYear, startMonth, 1);
            DateTime endDate = new(endYear, endMonth, 1);

            while (currentDate <= endDate)
            {
                DataRow newRow = forecastTable.NewRow();
                newRow["Year"] = currentDate.Year;
                newRow["Month"] = currentDate.ToString("MMM", System.Globalization.CultureInfo.InvariantCulture);

                // For every market column, apply the logic
                foreach (int colIndex in marketColIndices)
                {
                    double forecastValue = 0;

                    // 1. Get the "Anchor" (The previous month's value)
                    // If we have it in our generated forecastTable, use that (Rolling). 
                    // If not, look in historyMap (Start).
                    double lastActual = GetValue(currentDate.AddMonths(-1), colIndex, historyMap, forecastTable);

                    if (lastActual > 0)
                    {
                        // 2. Calculate Growth Ratios (MoM)
                        // Case 1: Look at Last Year (Same Month / Prev Month)
                        double ratioSum = 0;
                        int ratioCount = 0;

                        // Look back up to 5 years
                        for (int lookback = 1; lookback <= 5; lookback++)
                        {
                            DateTime pastTarget = currentDate.AddYears(-lookback); // e.g., Jan 25
                            DateTime pastPrev = pastTarget.AddMonths(-1);          // e.g., Dec 24

                            double pastTargetVal = GetValue(pastTarget, colIndex, historyMap, forecastTable);
                            double pastPrevVal = GetValue(pastPrev, colIndex, historyMap, forecastTable);

                            if (pastTargetVal > 0 && pastPrevVal > 0)
                            {
                                ratioSum += (pastTargetVal / pastPrevVal);
                                ratioCount++;
                            }
                        }

                        // 3. Apply Average Ratio
                        if (ratioCount > 0)
                        {
                            double avgRatio = ratioSum / ratioCount;
                            forecastValue = lastActual * avgRatio;
                        }
                    }

                    newRow[colIndex] = Math.Round(forecastValue, 0); // Round to nearest integer for orders
                }

                forecastTable.Rows.Add(newRow);

                // Add this new forecasted row to historyMap temporarily so the *next* iteration can find it (Rolling)
                // We create a temporary mapping just for the calculation loop
                // Note: We don't modify the original 'historicalData'
                historyMap[(currentDate.Year, currentDate.Month)] = newRow;

                currentDate = currentDate.AddMonths(1);
            }

            return forecastTable;
        }

        // Helper to find data in either History or the new Forecast table
        private static double GetValue(DateTime date, int colIndex, Dictionary<(int, int), DataRow> history, DataTable forecast)
        {
            // 1. Try History/Forecast Map (Since we add forecast rows to the map as we go, this covers both)
            if (history.ContainsKey((date.Year, date.Month)))
            {
                var row = history[(date.Year, date.Month)];
                bool v = double.TryParse(row[colIndex].ToString(), out double val);
                return val;
            }

            ArgumentNullException.ThrowIfNull(forecast);
            return 0;
        }
    }
}