using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace oPenEfficiency.Excel.Features.Rounding
{
    /// <summary>
    /// COM automation add-in exposing the consistent rounding worksheet functions
    /// OEROUND / OEROUNDUP / OEROUNDDOWN / OEROUNDNEAR to Excel. Each call delegates to
    /// <see cref="OeRoundEngine"/>, which returns a consistently-rounded value.
    ///
    /// The class is COM-visible with a stable GUID/ProgId; <see cref="RegisterFunction"/>
    /// adds the "Programmable" marker so Excel lists it as an automation add-in.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId(ProgId)]
    [Guid("2F1E7C4A-6B3D-4E8A-9C21-7A4F0E5B9D31")]
    public class OeRoundFunctions
    {
        public const string ProgId = "oPenEfficiency.OeRound";

        public double OEROUND(double x, double n) => OeRoundEngine.Evaluate(x, n, OeRoundMode.Optimize);
        public double OEROUNDUP(double x, double n) => OeRoundEngine.Evaluate(x, n, OeRoundMode.Up);
        public double OEROUNDDOWN(double x, double n) => OeRoundEngine.Evaluate(x, n, OeRoundMode.Down);
        public double OEROUNDNEAR(double x, double n) => OeRoundEngine.Evaluate(x, n, OeRoundMode.Near);

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.CreateSubKey(ProgrammableKey(type));
            }
            catch { /* best effort */ }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(ProgrammableKey(type), false);
            }
            catch { /* best effort */ }
        }

        private static string ProgrammableKey(Type type)
            => $@"CLSID\{{{type.GUID.ToString().ToUpperInvariant()}}}\Programmable";
    }
}
