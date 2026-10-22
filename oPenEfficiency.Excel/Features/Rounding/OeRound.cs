using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Excel.Features.Rounding
{
    /// <summary>Rounding direction for an OEROUND-family function.</summary>
    public enum OeRoundMode
    {
        /// <summary>OEROUND: nearest multiple, but the engine may flip it to keep sums consistent.</summary>
        Optimize,
        /// <summary>OEROUNDUP: always away from zero.</summary>
        Up,
        /// <summary>OEROUNDDOWN: always toward zero.</summary>
        Down,
        /// <summary>OEROUNDNEAR: always to the nearest multiple.</summary>
        Near
    }

    /// <summary>
    /// Pure rounding mathematics for consistent rounding.
    /// No Excel dependency, so this part is fully deterministic and testable.
    ///
    /// OEROUND(x, n) rounds x to the nearest multiple of n (n is the precision/unit,
    /// e.g. 1, 0.01, 5, 100).
    /// </summary>
    public static class OeRound
    {
        /// <summary>Rounds a single value to a multiple of <paramref name="n"/> in the given direction.</summary>
        public static double RoundMultiple(double x, double n, OeRoundMode mode)
        {
            if (n == 0 || double.IsNaN(n) || double.IsInfinity(n)) return x;

            double unit = Math.Abs(n);
            double q = x / unit;

            double k;
            switch (mode)
            {
                case OeRoundMode.Up:   // away from zero
                    k = x >= 0 ? Math.Ceiling(q) : Math.Floor(q);
                    break;
                case OeRoundMode.Down: // toward zero
                    k = x >= 0 ? Math.Floor(q) : Math.Ceiling(q);
                    break;
                default: // Near / Optimize base case
                    k = Math.Round(q, MidpointRounding.AwayFromZero);
                    break;
            }
            return k * unit;
        }

        /// <summary>
        /// Consistent rounding of a set of summands to multiples of <paramref name="n"/> so that
        /// the sum of the rounded values equals the nearest-rounded value of the exact sum.
        /// Uses the largest-remainder method: most values round to the nearest multiple, and only
        /// as few as necessary are flipped the other way, minimizing total deviation.
        /// </summary>
        /// <param name="values">The exact summand values.</param>
        /// <param name="n">The precision (unit) to round to.</param>
        /// <returns>The consistently rounded values, in the same order.</returns>
        public static double[] SolveConsistentSum(IReadOnlyList<double> values, double n)
        {
            int count = values.Count;
            var result = new double[count];
            if (count == 0) return result;

            double unit = Math.Abs(n);
            if (unit == 0 || double.IsNaN(unit) || double.IsInfinity(unit))
            {
                for (int i = 0; i < count; i++) result[i] = values[i];
                return result;
            }

            // Work in units of n.
            var floors = new long[count];
            var remainders = new double[count];
            double sumUnits = 0;
            for (int i = 0; i < count; i++)
            {
                double u = values[i] / unit;
                sumUnits += u;
                double f = Math.Floor(u);
                floors[i] = (long)f;
                remainders[i] = u - f; // in [0, 1)
            }

            // Target integer count for the whole group = nearest integer of the exact sum (in units).
            long target = (long)Math.Round(sumUnits, MidpointRounding.AwayFromZero);
            long baseSum = floors.Sum();
            long delta = target - baseSum; // how many +1 adjustments are needed (may be negative)

            var adjust = new int[count];

            if (delta > 0)
            {
                // Round up the 'delta' values with the largest remainders (closest to the next multiple).
                foreach (int i in Enumerable.Range(0, count)
                             .OrderByDescending(i => remainders[i])
                             .ThenByDescending(i => Math.Abs(values[i]))
                             .Take((int)Math.Min(delta, count)))
                {
                    adjust[i] = 1;
                }
            }
            else if (delta < 0)
            {
                // Round down the values with the smallest remainders.
                foreach (int i in Enumerable.Range(0, count)
                             .OrderBy(i => remainders[i])
                             .ThenByDescending(i => Math.Abs(values[i]))
                             .Take((int)Math.Min(-delta, count)))
                {
                    adjust[i] = -1;
                }
            }

            for (int i = 0; i < count; i++)
            {
                result[i] = (floors[i] + adjust[i]) * unit;
            }
            return result;
        }
    }
}
