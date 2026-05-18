namespace oPenEfficiency.Utils
{
    /// <summary>
    /// Centralized color constants used across the oPenEfficiency add-in.
    /// All colors are in BGR format (0xBBGGRR) for PowerPoint COM interop.
    /// </summary>
    public static class ColorConstants
    {
        // Basic Colors
        public const int Black = 0x000000;
        public const int White = 0xFFFFFF;
        public const int Red = 0x0000FF;
        public const int Green = 0x00FF00;
        public const int Blue = 0xFF0000;
        public const int Yellow = 0x00FFFF;
        public const int Cyan = 0xFFFF00;
        public const int Magenta = 0xFF00FF;
        public const int Gray = 0x808080;
        public const int LightGray = 0xD3D3D3;
        public const int DarkGray = 0x404040;

        // Material Design Blues
        public const int Blue50 = 0xE3F2FD;
        public const int Blue100 = 0xBBDEFB;
        public const int Blue200 = 0x90CAF9;
        public const int Blue300 = 0x64B5F6;
        public const int Blue400 = 0x42A5F5;
        public const int Blue500 = 0x2196F3;
        public const int Blue600 = 0x1E88E5;
        public const int Blue700 = 0x1976D2;
        public const int Blue800 = 0x1565C0;
        public const int Blue900 = 0x0D47A1;

        // Material Design Greens
        public const int Green50 = 0xE8F5E9;
        public const int Green100 = 0xC8E6C9;
        public const int Green200 = 0xA5D6A7;
        public const int Green300 = 0x81C784;
        public const int Green400 = 0x66BB6A;
        public const int Green500 = 0x4CAF50;
        public const int Green600 = 0x43A047;
        public const int Green700 = 0x388E3C;
        public const int Green800 = 0x2E7D32;
        public const int Green900 = 0x1B5E20;

        // Material Design Reds
        public const int Red50 = 0xFFEBEE;
        public const int Red100 = 0xFFCDD2;
        public const int Red200 = 0xEF9A9A;
        public const int Red300 = 0xE57373;
        public const int Red400 = 0xEF5350;
        public const int Red500 = 0xF44336;
        public const int Red600 = 0xE53935;
        public const int Red700 = 0xD32F2F;
        public const int Red800 = 0xC62828;
        public const int Red900 = 0xB71C1C;

        // Material Design Oranges
        public const int Orange50 = 0xFFF3E0;
        public const int Orange100 = 0xFFE0B2;
        public const int Orange200 = 0xFFCC80;
        public const int Orange300 = 0xFFB74D;
        public const int Orange400 = 0xFFA726;
        public const int Orange500 = 0xFF9800;
        public const int Orange600 = 0xFB8C00;
        public const int Orange700 = 0xF57C00;
        public const int Orange800 = 0xEF6C00;
        public const int Orange900 = 0xE65100;

        // Sticker Colors (from AddStickerFeature)
        public const int StickerGreen = 0x50AF4C;         // BGR: #4CAF50
        public const int StickerWhite = 0xFFFFFF;

        // Thermometer Colors (from ThermometerFeature)
        public const int ThermometerLightGray = 0xEEEEEE;
        public const int ThermometerDarkGray = 0x808080;



        // Chart Colors - Harvey Ball
        public const int HarveyBallStroke = 0x000000;     // BGR: Black

        // Chart Colors - Star Rating
        public const int StarRatingDim = 0xD0D0D0;        // BGR: Light grey for empty stars
        public const int StarRatingStroke = 0x808080;     // BGR: Grey stroke

        // Chart Colors - Traffic Light (updated from TrafficLightFeature)
        public const int TrafficLightRed = 0x3232E8;      // BGR: #E83232
        public const int TrafficLightYellow = 0x17C8F5;   // BGR: #F5C817
        public const int TrafficLightGreen = 0x3EAF4C;    // BGR: #4CAF3E



        // UI Colors
        public const int SelectionHighlight = 0xFF9900;   // Orange in BGR
        public const int DisabledGray = 0xB0B0B0;

        /// <summary>
        /// Converts a hex string (#RRGGBB or RRGGBB) to BGR integer format for PowerPoint.
        /// </summary>
        public static int FromHex(string hex)
        {
            if (string.IsNullOrEmpty(hex)) return Black;

            hex = hex.Replace("#", "").Trim();
            if (hex.Length != 6) return Black;

            int r = System.Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = System.Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = System.Convert.ToInt32(hex.Substring(4, 2), 16);

            return r | (g << 8) | (b << 16);
        }

        /// <summary>
        /// Converts RGB values to BGR integer format for PowerPoint.
        /// </summary>
        public static int FromRgb(int r, int g, int b)
        {
            return r | (g << 8) | (b << 16);
        }

        /// <summary>
        /// Converts a BGR integer to hex string format (#RRGGBB).
        /// </summary>
        public static string ToHex(int bgr)
        {
            int r = bgr & 0xFF;
            int g = (bgr >> 8) & 0xFF;
            int b = (bgr >> 16) & 0xFF;
            return $"#{r:X2}{g:X2}{b:X2}";
        }

        // Dimension Constants - Chart Defaults

        // Dimension Constants - Sticker Defaults
        public const float StickerWidth = 100f;
        public const float StickerHeight = 30f;
        public const float StickerMargin = 10f;

        // Dimension Constants - Spacing Defaults
        public const float SpacingDefaultIncrement = 5f;


        // Dimension Constants - Comparison Tolerances
        public const float FillComparisonTolerance = 0.1f;
        public const float GradientStopPositionTolerance = 0.01f;
        public const float SizeComparisonTolerance = 0.1f;
    }
}
