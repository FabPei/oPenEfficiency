using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using oPenEfficiency.Utils;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Excel.Features.Rounding
{
    /// <summary>
    /// Global engine backing the live OEROUND worksheet functions. It captures each
    /// OEROUND cell's exact input during a recalculation, then (after the pass) solves
    /// every rounding group consistently and stores per-cell overrides that the
    /// functions return on the next recalculation. A reentrancy guard keeps the forced
    /// recalculation from looping.
    ///
    /// Groups are the ranges the user wrapped, stored as workbook-scoped defined names
    /// with the <see cref="GroupNamePrefix"/> prefix.
    /// </summary>
    public static class OeRoundEngine
    {
        public const string GroupNamePrefix = "oPE_OER_";

        private static readonly object _lock = new object();
        private static Xl.Application _app;
        private static bool _hooked;
        private static bool _applyingOverrides;

        // Exact input captured for each OEROUND cell during the current calc pass.
        private static readonly Dictionary<string, CellInput> _inputs = new Dictionary<string, CellInput>();
        // Consistent override to return for a cell, valid only while its exact input matches.
        private static readonly Dictionary<string, CellOverride> _overrides = new Dictionary<string, CellOverride>();

        private struct CellInput { public double X; public double N; public OeRoundMode Mode; }
        private struct CellOverride { public double X; public double Value; }

        public static void Initialize(Xl.Application app)
        {
            lock (_lock)
            {
                _app = app;
                if (_app != null && !_hooked)
                {
                    _app.AfterCalculate += OnAfterCalculate;
                    _hooked = true;
                }
            }
        }

        /// <summary>
        /// Ensures the engine has an Excel Application to work with. The worksheet
        /// functions may run in a different AppDomain than the VSTO add-in, so we
        /// fall back to the running instance and hook AfterCalculate lazily.
        /// </summary>
        private static void EnsureApp()
        {
            if (_app != null) return;
            try
            {
                var app = (Xl.Application)Marshal.GetActiveObject("Excel.Application");
                Initialize(app);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "OeRoundEngine.EnsureApp");
            }
        }

        public static void Shutdown()
        {
            lock (_lock)
            {
                try { if (_app != null && _hooked) _app.AfterCalculate -= OnAfterCalculate; }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "OeRoundEngine.Detach.AfterCalculate"); }
                _hooked = false;
                _app = null;
                _inputs.Clear();
                _overrides.Clear();
            }
        }

        /// <summary>Entry point used by the worksheet functions. Records the input and returns the value to display.</summary>
        public static double Evaluate(double x, double n, OeRoundMode mode)
        {
            try
            {
                EnsureApp();
                string key = CallerKey();
                double naive = OeRound.RoundMultiple(x, n, mode);

                if (key == null) return naive;

                lock (_lock)
                {
                    _inputs[key] = new CellInput { X = x, N = n, Mode = mode };

                    if (mode == OeRoundMode.Optimize &&
                        _overrides.TryGetValue(key, out var ov) &&
                        NearlyEqual(ov.X, x))
                    {
                        return ov.Value;
                    }
                }

                return naive;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "OeRoundEngine.Evaluate");
                return OeRound.RoundMultiple(x, n, mode);
            }
        }

        private static void OnAfterCalculate()
        {
            lock (_lock)
            {
                if (_applyingOverrides || _app == null) return;
            }

            try
            {
                bool changed = SolveAllGroups();
                if (changed)
                {
                    lock (_lock) { _applyingOverrides = true; }
                    try { _app.Calculate(); }
                    finally { lock (_lock) { _applyingOverrides = false; } }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "OeRoundEngine.OnAfterCalculate");
            }
        }

        /// <summary>Solves every registered group; returns true if any override changed.</summary>
        private static bool SolveAllGroups()
        {
            bool changed = false;
            Xl.Workbook wb = null;
            try { wb = _app.ActiveWorkbook; }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "OeRoundEngine.SolveAllGroups.ActiveWorkbook"); }
            if (wb == null) return false;

            foreach (Xl.Name name in wb.Names)
            {
                string nm = null;
                try { nm = name.Name; }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "OeRoundEngine.SolveAllGroups.Name"); }
                if (nm == null || !nm.Contains(GroupNamePrefix)) continue;

                Xl.Range range = null;
                try { range = name.RefersToRange; }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "OeRoundEngine.SolveAllGroups.Range"); }
                if (range == null) continue;

                changed |= SolveGroup(range);
            }
            return changed;
        }

        private static bool SolveGroup(Xl.Range range)
        {
            var keys = new List<string>();
            var values = new List<double>();
            double n = 0;

            foreach (Xl.Range cell in range.Cells)
            {
                string key = RangeKey(cell);
                if (key == null) continue;

                lock (_lock)
                {
                    if (!_inputs.TryGetValue(key, out var input)) continue;
                    if (input.Mode != OeRoundMode.Optimize) continue; // forced cells stay fixed
                    keys.Add(key);
                    values.Add(input.X);
                    n = input.N;
                }
            }

            if (keys.Count == 0 || n == 0) return false;

            double[] solved = OeRound.SolveConsistentSum(values, n);

            bool changed = false;
            lock (_lock)
            {
                for (int i = 0; i < keys.Count; i++)
                {
                    var newOv = new CellOverride { X = values[i], Value = solved[i] };
                    if (!_overrides.TryGetValue(keys[i], out var old) ||
                        !NearlyEqual(old.Value, newOv.Value) || !NearlyEqual(old.X, newOv.X))
                    {
                        _overrides[keys[i]] = newOv;
                        changed = true;
                    }
                }
            }
            return changed;
        }

        /// <summary>True if the cell was rounded away from its nearest multiple to keep the sum consistent.</summary>
        public static bool IsCheated(Xl.Range cell)
        {
            try
            {
                string key = RangeKey(cell);
                if (key == null) return false;
                lock (_lock)
                {
                    if (_inputs.TryGetValue(key, out var input) &&
                        _overrides.TryGetValue(key, out var ov) &&
                        NearlyEqual(ov.X, input.X))
                    {
                        double naive = OeRound.RoundMultiple(input.X, input.N, input.Mode);
                        return !NearlyEqual(naive, ov.Value);
                    }
                }
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "OeRoundEngine.IsCheated"); }
            return false;
        }

        private static string CallerKey()
        {
            try
            {
                var caller = _app?.Caller as Xl.Range;
                return caller != null ? RangeKey(caller) : null;
            }
            catch
            {
                return null;
            }
        }

        private static string RangeKey(Xl.Range cell)
        {
            try
            {
                var ws = cell.Worksheet;
                var wb = (Xl.Workbook)ws.Parent;
                return wb.Name + "|" + ws.Name + "|" + cell.Address;
            }
            catch
            {
                return null;
            }
        }

        private static bool NearlyEqual(double a, double b)
        {
            return Math.Abs(a - b) <= 1e-9 * Math.Max(1.0, Math.Max(Math.Abs(a), Math.Abs(b)));
        }
    }
}
