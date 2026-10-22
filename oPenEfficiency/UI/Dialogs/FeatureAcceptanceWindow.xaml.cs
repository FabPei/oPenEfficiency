using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.UI.Dialogs
{
    /// <summary>
    /// Dev/QA window for the feature acceptance harness. Every command launches the existing
    /// eng/ infrastructure (gates, evidence binding, scenario runners) in a visible window;
    /// this class holds no gate logic of its own. Available only where the repository checkout
    /// is reachable (FeatureAcceptanceStatusService.IsAvailable).
    /// </summary>
    public partial class FeatureAcceptanceWindow : Window
    {
        private List<FeatureAcceptanceScenarioStatus> _rows = new List<FeatureAcceptanceScenarioStatus>();
        private string _repositoryRoot;

        public FeatureAcceptanceWindow()
        {
            InitializeComponent();
            Reload();
        }

        /// <summary>Row count for the host acceptance probe (qa-acceptance-window smoke).</summary>
        public int RowCount
        {
            get { return _rows.Count; }
        }

        private void Reload()
        {
            try
            {
                _rows = FeatureAcceptanceStatusService.LoadStatuses(out _repositoryRoot);
                RepositoryLabel.Text = _repositoryRoot == null
                    ? "repository: (not found) - the QA window needs a checkout"
                    : "repository: " + _repositoryRoot;
                ApplyView();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FeatureAcceptanceWindow.Reload");
                _rows = new List<FeatureAcceptanceScenarioStatus>();
                RepositoryLabel.Text = "repository: status load failed (see log)";
                ApplyView();
            }
        }

        private void ApplyView()
        {
            IEnumerable<FeatureAcceptanceScenarioStatus> view = _rows;

            var severityChoice = SeverityFilter.SelectedIndex;
            if (severityChoice == 1)
                view = view.Where(r => r.Severity != "none");
            else if (severityChoice >= 2)
            {
                var wanted = (SeverityFilter.SelectedItem as ComboBoxItem).Content.ToString();
                view = view.Where(r => r.Severity == wanted);
            }

            var sortChoice = SortSelector.SelectedIndex;
            if (sortChoice == 1)
                view = view.OrderByDescending(r => r.EvidenceCompletedUtc ?? DateTime.MinValue);
            else if (sortChoice == 2)
                view = view.OrderBy(r => r.ScenarioId, StringComparer.OrdinalIgnoreCase);
            else
            {
                var order = new Dictionary<string, int>(StringComparer.Ordinal)
                {
                    { "failed", 0 }, { "missing-evidence", 1 }, { "stale", 2 }, { "none", 3 }
                };
                view = view.OrderBy(r => order[r.Severity]).ThenBy(r => r.ScenarioId, StringComparer.OrdinalIgnoreCase);
            }

            var materialized = view.ToList();
            RowsGrid.ItemsSource = materialized;
            IssueCountLabel.Text = string.Format(
                "{0} shown | {1} issues | {2} green",
                materialized.Count,
                materialized.Count(r => r.Severity != "none"),
                _rows.Count(r => r.EvidenceState == "Green"));
        }

        private void OnRefresh(object sender, RoutedEventArgs e)
        {
            Reload();
        }

        private void OnFilterChanged(object sender, SelectionChangedEventArgs e)
        {
            if (RowsGrid != null && SeverityFilter != null && SortSelector != null && IsLoaded)
                ApplyView();
        }

        private void LaunchVisibleProcess(string fileName, string arguments)
        {
            if (string.IsNullOrWhiteSpace(_repositoryRoot))
            {
                MessageBox.Show(this, "Repository checkout not found - this command needs the dev environment.",
                    "openEfficiency QA", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                WorkingDirectory = _repositoryRoot,
                UseShellExecute = true
            });
        }

        private void OnRunGate(object sender, RoutedEventArgs e)
        {
            // The window only orchestrates: Test-Changed stays the single gate implementation.
            LaunchVisibleProcess("powershell.exe",
                "-NoProfile -ExecutionPolicy Bypass -File \"" + System.IO.Path.Combine(_repositoryRoot ?? string.Empty, "eng", "Test-Changed.ps1") + "\"");
        }

        private void OnRebindEvidence(object sender, RoutedEventArgs e)
        {
            LaunchVisibleProcess("powershell.exe",
                "-NoProfile -ExecutionPolicy Bypass -File \"" + System.IO.Path.Combine(_repositoryRoot ?? string.Empty, "eng", "Bind-OfficeEvidenceAndReviews.ps1") + "\"");
        }

        private void OnRerunFeature(object sender, RoutedEventArgs e)
        {
            var selected = RowsGrid.SelectedItem as FeatureAcceptanceScenarioStatus;
            if (selected == null || !selected.HasRunner || string.IsNullOrWhiteSpace(_repositoryRoot))
            {
                MessageBox.Show(this,
                    "Select a scenario row with a runner (office/core scenarios with a manifest runner entry).",
                    "openEfficiency QA", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var script = System.IO.Path.Combine(_repositoryRoot, selected.RunnerPath.Replace('/', System.IO.Path.DirectorySeparatorChar));
            var arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\" " + selected.RunnerArguments;
            LaunchVisibleProcess("powershell.exe", arguments);
        }

        private void OnOpenEvidenceFolder(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_repositoryRoot)) return;
            var folder = System.IO.Path.Combine(_repositoryRoot, "TestResults");
            if (!System.IO.Directory.Exists(folder)) System.IO.Directory.CreateDirectory(folder);
            LaunchVisibleProcess(folder, string.Empty);
        }

        private void OnOpenChecklist(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_repositoryRoot)) return;
            var checklist = System.IO.Path.Combine(_repositoryRoot, "docs", "product", "qa-001-feature-acceptance-harness-todo.md");
            if (System.IO.File.Exists(checklist)) LaunchVisibleProcess(checklist, string.Empty);
        }

        private void OnCopySummary(object sender, RoutedEventArgs e)
        {
            try
            {
                var summary = FeatureAcceptanceStatusService.BuildSummary(_rows, _repositoryRoot);
                Clipboard.SetText(summary);
                IssueCountLabel.Text = "summary copied to clipboard (" + summary.Length + " chars)";
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FeatureAcceptanceWindow.OnCopySummary");
            }
        }

        private void OnRowDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OnRerunFeature(sender, e);
        }
    }
}
