using System;

namespace oPenEfficiency.Models
{
    public class SeriesConfig
    {
        public string SeriesId { get; set; }
        public string SeriesType { get; set; } // "Numbers", "Letters", "Days", "Months", "Quarters", "Years"
        public string Language { get; set; }
        public string StartValue { get; set; }
        public string EndValue { get; set; }
        public int Count { get; set; }
        public int Step { get; set; }
        public string FormatString { get; set; }
        public string Direction { get; set; } // "Horizontal", "Vertical", "Matrix"
        public int MatrixColumns { get; set; }
        public double SpacingX { get; set; }
        public double SpacingY { get; set; }
        public bool AutoFitToBounds { get; set; }
        public string BoundaryType { get; set; } // "Slide", "SelectedShape"
    }
}
