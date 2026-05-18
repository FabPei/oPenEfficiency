using System;

namespace oPenEfficiency.Utils
{
    public static class ColorMath
    {
        // Convert BGR (PowerPoint RGB integer) and Hex to CIELAB and calculate Delta-E
        public static double CalculateDeltaE(int bgr1, string hex2)
        {
            var lab1 = RgbToLab(BgrToRgb(bgr1));
            var lab2 = RgbToLab(HexToRgb(hex2));
            
            return Math.Sqrt(Math.Pow(lab2.L - lab1.L, 2) + Math.Pow(lab2.a - lab1.a, 2) + Math.Pow(lab2.b - lab1.b, 2));
        }

        public static (int R, int G, int B) BgrToRgb(int bgr)
        {
            return (bgr & 0xFF, (bgr & 0xFF00) >> 8, (bgr & 0xFF0000) >> 16);
        }

        public static (int R, int G, int B) HexToRgb(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return (0, 0, 0);
            hex = hex.Replace("#", "");
            if (hex.Length == 6)
            {
                try {
                    return (
                        Convert.ToInt32(hex.Substring(0, 2), 16),
                        Convert.ToInt32(hex.Substring(2, 2), 16),
                        Convert.ToInt32(hex.Substring(4, 2), 16)
                    );
                } catch { }
            }
            return (0, 0, 0);
        }

        public static (double L, double a, double b) RgbToLab((int R, int G, int B) rgb)
        {
            // RGB to XYZ
            double r = rgb.R / 255.0;
            double g = rgb.G / 255.0;
            double b = rgb.B / 255.0;

            r = (r > 0.04045) ? Math.Pow((r + 0.055) / 1.055, 2.4) : r / 12.92;
            g = (g > 0.04045) ? Math.Pow((g + 0.055) / 1.055, 2.4) : g / 12.92;
            b = (b > 0.04045) ? Math.Pow((b + 0.055) / 1.055, 2.4) : b / 12.92;

            r *= 100.0;
            g *= 100.0;
            b *= 100.0;

            double x = r * 0.4124 + g * 0.3576 + b * 0.1805;
            double y = r * 0.2126 + g * 0.7152 + b * 0.0722;
            double z = r * 0.0193 + g * 0.1192 + b * 0.9505;

            // XYZ to Lab
            x /= 95.047;
            y /= 100.000;
            z /= 108.883;

            x = (x > 0.008856) ? Math.Pow(x, 1.0 / 3.0) : (7.787 * x) + (16.0 / 116.0);
            y = (y > 0.008856) ? Math.Pow(y, 1.0 / 3.0) : (7.787 * y) + (16.0 / 116.0);
            z = (z > 0.008856) ? Math.Pow(z, 1.0 / 3.0) : (7.787 * z) + (16.0 / 116.0);

            return ((116.0 * y) - 16.0, 500.0 * (x - y), 200.0 * (y - z));
        }

        public static string FindClosestColor(int bgr, System.Collections.Generic.IEnumerable<string> allowedColors)
        {
            string closest = null;
            double minDelta = double.MaxValue;

            foreach (var hex in allowedColors)
            {
                double delta = CalculateDeltaE(bgr, hex);
                if (delta < minDelta)
                {
                    minDelta = delta;
                    closest = hex;
                }
            }

            return closest;
        }
    }
}