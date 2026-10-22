using System;
using System.Runtime.InteropServices;

namespace oPenEfficiency.Utils
{
    /// <summary>
    /// Narrow helpers for expected COM probes and cleanup that must not interrupt shutdown.
    /// Unexpected failures are never silently converted into an expected COM result.
    /// </summary>
    public static class ExceptionPolicy
    {
        public static bool IsExpectedComFailure(Exception exception)
        {
            return exception is COMException || exception is InvalidComObjectException;
        }

        public static bool TryComProbe<T>(Func<T> operation, out T value)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));

            try
            {
                value = operation();
                return true;
            }
            catch (Exception ex) when (IsExpectedComFailure(ex))
            {
                value = default(T);
                return false;
            }
        }

        public static void RunBestEffortComCleanup(Action cleanup, string context)
        {
            RunBestEffort(
                cleanup,
                context,
                IsExpectedComFailure,
                ExceptionLogger.Log);
        }

        public static void ReportOptionalComFailure(Exception exception, string context)
        {
            if (exception == null) throw new ArgumentNullException(nameof(exception));
            if (string.IsNullOrWhiteSpace(context))
                throw new ArgumentException("A diagnostic context is required.", nameof(context));

            if (!IsExpectedComFailure(exception))
            {
                ExceptionLogger.Log(exception, context);
            }
        }

        public static void RunBestEffort(
            Action operation,
            string context,
            Func<Exception, bool> isExpected,
            Action<Exception, string> logUnexpected)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));
            if (string.IsNullOrWhiteSpace(context))
                throw new ArgumentException("A diagnostic context is required.", nameof(context));
            if (isExpected == null) throw new ArgumentNullException(nameof(isExpected));
            if (logUnexpected == null) throw new ArgumentNullException(nameof(logUnexpected));

            try
            {
                operation();
            }
            catch (Exception ex) when (isExpected(ex))
            {
                // Expected absence, stale COM state, or already-released cleanup target.
            }
            catch (Exception ex)
            {
                logUnexpected(ex, context);
            }
        }
    }
}
