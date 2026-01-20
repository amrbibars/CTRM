using System;
using System.Collections.Generic;

namespace CTRM.Models
{
    public sealed class WeeklyMatrix
    {
        public string SourcePath { get; init; } = "";
        public string SheetName { get; init; } = "";
        public IReadOnlyList<string> Headers { get; init; } = Array.Empty<string>(); // e.g., Mon..Sun
        public IReadOnlyList<IntervalRow> Rows { get; init; } = Array.Empty<IntervalRow>();
        public bool HasTotalsRow { get; init; }
        public IReadOnlyList<double>? TotalsPerColumn { get; init; }  // optional, if present
        public int? IntervalMinutes { get; init; }                     // you’ll fill later based on user input
    }

    public sealed class IntervalRow
    {
        // Optional label like "08:00", keep null if not present
        public string? Label { get; init; }
        // Same order as Headers
        public IReadOnlyList<double> Values { get; init; } = Array.Empty<double>();
    }
}
