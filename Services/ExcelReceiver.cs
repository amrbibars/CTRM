using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using CTRM.Models;
using System.Data;  // Added for importing Orders in Budgeting Window

namespace CTRM.Services
{
    public static class ExcelReceiver
    {
        /// <summary>
        /// Reads a matrix (volumes or AHT) from an Excel sheet.
        /// Expected layout:
        ///   Row 1: headers (weekdays etc.)
        ///   Rows 2..N-1: interval rows (24..48 typical, but not enforced here)
        ///   Optional last row: totals (detected by first cell text startswith "total")
        /// First column may be interval labels (e.g., "08:00"); if non-numeric, we treat it as label column.
        /// </summary>
        public static WeeklyMatrix LoadMatrix(string path, string? sheetName = null)
        {
            using var wb = new XLWorkbook(path);
            var ws = sheetName is null ? wb.Worksheets.First() :
                                         wb.Worksheets.Worksheet(sheetName);

            var used = ws.RangeUsed() ?? throw new InvalidOperationException("No used cells found.");
            var firstRow = used.FirstRowUsed();
            var firstCol = used.FirstColumnUsed();

            int r0 = firstRow.RowNumber();
            int c0 = firstCol.ColumnNumber();
            int r1 = used.LastRow().RowNumber();
            int c1 = used.LastColumn().ColumnNumber();

            // Read header row (r0)
            var headers = new List<string>();
            for (int c = c0; c <= c1; c++)
            {
                var text = ws.Cell(r0, c).GetString().Trim();
                headers.Add(text);
            }

            // Detect whether first data column is a label column
            // Heuristic: if the first column (below header) is non-numeric in majority, treat it as labels
            bool firstColSeemsLabel = false;
            int sampleCount = Math.Min(10, Math.Max(0, r1 - r0));
            int nonNumeric = 0;
            for (int r = r0 + 1; r <= r0 + sampleCount; r++)
            {
                var s = ws.Cell(r, c0).GetString();
                if (!TryCellToDouble(ws.Cell(r, c0), out _)) nonNumeric++;
            }
            firstColSeemsLabel = nonNumeric > sampleCount / 2;

            // If label column exists, data starts at cDataStart; else data starts at c0
            int cDataStart = firstColSeemsLabel ? c0 + 1 : c0;
            var effectiveHeaders = headers.Skip(firstColSeemsLabel ? 1 : 0).Select(h => h.Trim()).ToList();

            if (effectiveHeaders.Count == 0)
                throw new InvalidOperationException("No data headers detected.");

            // Read body rows
            var rows = new List<IntervalRow>();
            bool hasTotals = false;
            List<double>? totals = null;

            for (int r = r0 + 1; r <= r1; r++)
            {
                // detect totals row (look at first non-empty cell in the row)
                var firstCellText = ws.Cell(r, c0).GetString().Trim().ToLowerInvariant();
                bool looksLikeTotals = firstCellText.Contains("total") ||firstCellText.Contains("overall");

                var values = new List<double>(capacity: effectiveHeaders.Count);
                bool anyValuePresent = false;

                for (int c = 0; c < effectiveHeaders.Count; c++)
                {
                    var cell = ws.Cell(r, cDataStart + c);
                    if (TryCellToDouble(cell, out double v))
                    {
                        values.Add(v);
                        anyValuePresent = true;
                    }
                    else
                    {
                        // empty or text → treat as 0
                        values.Add(0d);
                    }
                }

                if (!anyValuePresent)
                {
                    // blank row → skip
                    continue;
                }

                if (looksLikeTotals)
                {
                    hasTotals = true;
                    totals = values;
                    // do not add as interval row
                }
                else
                {
                    string? label = firstColSeemsLabel ? ws.Cell(r, c0).GetString().Trim() : null;
                    rows.Add(new IntervalRow { Label = label, Values = values });
                }
            }

            return new WeeklyMatrix
            {
                SourcePath = path,
                SheetName = ws.Name,
                Headers = effectiveHeaders,
                Rows = rows,
                HasTotalsRow = hasTotals,
                TotalsPerColumn = totals,
                IntervalMinutes = null // you’ll plug this from UI later
            };
        }

        private static bool TryCellToDouble(IXLCell cell, out double value)
        {
            // handles numbers and numeric strings, uses invariant culture
            if (cell.DataType == XLDataType.Number)
            {
                value = cell.GetDouble();
                return true;
            }
            var s = cell.GetString().Trim();
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return true;

            value = 0;
            return false;
        }

        /// <summary>
        /// Imports Market Orders.
        /// Logic:
        /// 1. Reads Year (Col 1) and Month (Col 2).
        /// 2. validates (Year + Month) is unique.
        /// 3. Reads Markets dynamically from Col 3 until "Grand Total".
        /// 4. Excludes the bottom "Grand Total" row.
        /// </summary>
        public static DataTable ImportOrders(string path)
        {
            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheets.First();
            var usedRange = ws.RangeUsed() ?? throw new Exception("The selected Excel sheet is empty. Please provide a file with order data.");
            var rows = usedRange.RowsUsed().ToList();

            if (rows.Count < 2) throw new Exception("File is empty or missing headers.");

            DataTable dt = new();

            // --- 1. Parse Headers (First Row) ---
            var headerRow = rows[0];
            int totalColumnIndex = -1;

            // We expect: Col 1 = Year, Col 2 = Month, Col 3+ = Markets
            dt.Columns.Add("Year", typeof(int));
            dt.Columns.Add("Month", typeof(string));

            // Scan for markets starting at column 3 (index 2 in 0-based, or cell 3)
            // Note: ClosedXML cells are 1-based.
            int colCount = headerRow.CellCount();

            for (int c = 3; c <= colCount; c++)
            {
                string headerVal = headerRow.Cell(c).GetString().Trim();

                // Stop if we hit the Total column
                if (headerVal.Equals("Grand Total", StringComparison.OrdinalIgnoreCase) ||
                    headerVal.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    totalColumnIndex = c;
                    break;
                }

                // Add Market Column (using Double for order counts)
                dt.Columns.Add(headerVal, typeof(double));
            }

            // Validate we found markets
            if (dt.Columns.Count < 3)
                throw new Exception("No market columns found before 'Grand Total'.");

            // --- 2. Parse Data Rows ---
            // Track unique keys to prevent duplicates
            var uniqueKeys = new HashSet<(string, string)>();

            // Skip header (index 0), iterate data
            // Check if the last row is a total row and exclude it
            var lastRow = rows[^1];
            bool lastRowIsTotal = lastRow.Cell(1).GetString().Contains("Total", StringComparison.OrdinalIgnoreCase) ||
                                  lastRow.Cell(2).GetString().Contains("Total", StringComparison.OrdinalIgnoreCase);

            int rowLoopEnd = lastRowIsTotal ? rows.Count - 2 : rows.Count - 1;

            // Rows in 'rows' list are 0-indexed relative to the list
            for (int i = 1; i <= rowLoopEnd; i++)
            {
                IXLRangeRow row = rows[i];

                // 1. Parse Year and Month
                string yearText = row.Cell(1).GetString().Trim();
                string monthText = row.Cell(2).GetString().Trim();

                // Skip empty rows if any
                if (string.IsNullOrEmpty(yearText) && string.IsNullOrEmpty(monthText)) continue;

                // Validation: Duplicates [Constraint B]
                if (uniqueKeys.Contains((yearText, monthText)))
                {
                    throw new InvalidOperationException($"Duplicate data found for Year: {yearText}, Month: {monthText}. Please check the source file.");
                }
                uniqueKeys.Add((yearText, monthText));

                // Create DataRow
                DataRow dr = dt.NewRow();

                // Safe parse Year
                if (int.TryParse(yearText, out int year)) dr["Year"] = year;
                else dr["Year"] = 0; // Or throw error

                dr["Month"] = monthText;

                // 2. Parse Market Data
                // Match the columns we added to DataTable (skipping first 2: Year, Month)
                int dtColIdx = 2;

                // Loop through the same columns we identified in headers
                // Stop at totalColumnIndex (or end of row if no total found)
                int limit = totalColumnIndex == -1 ? row.CellCount() : totalColumnIndex - 1;

                for (int c = 3; c <= limit; c++)
                {
                    if (dtColIdx >= dt.Columns.Count) break;

                    // Parse value
                    if (TryCellToDouble(row.Cell(c), out double val))
                    {
                        dr[dtColIdx] = val;
                    }
                    else
                    {
                        dr[dtColIdx] = 0;
                    }
                    dtColIdx++;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

    }
}
