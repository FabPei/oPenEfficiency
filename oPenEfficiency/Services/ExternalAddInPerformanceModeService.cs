using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Services
{
    internal sealed class ExternalAddInPerformanceModeService : IDisposable
    {
        private static ExternalAddInPerformanceModeService _current;
        private readonly PowerPoint.Application _application;
        private readonly Dictionary<string, string> _pausedBySession = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private bool _enabled;

        public ExternalAddInPerformanceModeService(PowerPoint.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public static ExternalAddInPerformanceModeService For(PowerPoint.Application application)
        {
            if (_current == null) _current = new ExternalAddInPerformanceModeService(application);
            return _current;
        }

        public bool IsEnabled => _enabled;

        public IReadOnlyList<ExternalAddInState> Discover()
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
            var selected = new HashSet<string>(config.PerformanceModeTargetIds ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            var result = new List<ExternalAddInState>();
            Office.COMAddIns addIns = null;
            try
            {
                addIns = _application.COMAddIns;
                for (int index = 1; index <= addIns.Count; index++)
                {
                    Office.COMAddIn addIn = null;
                    try
                    {
                        addIn = addIns.Item(index);
                        string progId = addIn.ProgId ?? string.Empty;
                        if (ExternalAddInIdentity.IsOpenEfficiency(progId, addIn.Description)) continue;
                        bool connected = addIn.Connect;
                        result.Add(new ExternalAddInState
                        {
                            Id = progId,
                            DisplayName = string.IsNullOrWhiteSpace(addIn.Description) ? progId : addIn.Description,
                            IsConnected = connected,
                            IsSelected = selected.Contains(progId),
                            Status = _pausedBySession.ContainsKey(progId) ? "Paused" : connected ? "Active" : "Inactive",
                            RestorationCapability = _pausedBySession.ContainsKey(progId) ? "Pending verification" : "Not tested"
                        });
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "PerformanceMode.Discover.Item");
                    }
                    finally
                    {
                        if (addIn != null && Marshal.IsComObject(addIn)) Marshal.ReleaseComObject(addIn);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "PerformanceMode.Discover");
            }
            finally
            {
                if (addIns != null && Marshal.IsComObject(addIns)) Marshal.ReleaseComObject(addIns);
            }
            return result.OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList().AsReadOnly();
        }

        public bool SaveSelection(IEnumerable<string> ids)
        {
            var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
            config.PerformanceModeTargetIds = ExternalAddInIdentity.Normalize(ids).ToList();
            return JsonService.TrySave(config, "sidebar_config.json");
        }

        public bool SetEnabled(bool enabled)
        {
            if (enabled == _enabled) return true;
            bool success = enabled ? PauseSelected() : RestorePaused();
            if (enabled) _enabled = true;
            else if (_pausedBySession.Count == 0) _enabled = false;
            return success;
        }

        public void Dispose()
        {
            RestorePaused();
            _enabled = false;
            if (ReferenceEquals(_current, this)) _current = null;
        }

        private bool PauseSelected()
        {
            var targets = new HashSet<string>(
                (JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig()).PerformanceModeTargetIds ?? new List<string>(),
                StringComparer.OrdinalIgnoreCase);
            bool success = true;
            foreach (string id in targets)
            {
                if (ExternalAddInIdentity.IsOpenEfficiency(id, null)) continue;
                Office.COMAddIn addIn = null;
                Office.COMAddIns addIns = null;
                try
                {
                    addIns = _application.COMAddIns;
                    addIn = addIns.Item(id);
                    if (!addIn.Connect) continue;
                    string description = addIn.Description ?? id;
                    addIn.Connect = false;
                    if (addIn.Connect) { success = false; continue; }
                    _pausedBySession[id] = description;
                }
                catch (Exception ex)
                {
                    success = false;
                    ExceptionLogger.Log(ex, "PerformanceMode.Pause." + id);
                }
                finally
                {
                    if (addIn != null && Marshal.IsComObject(addIn)) Marshal.ReleaseComObject(addIn);
                    if (addIns != null && Marshal.IsComObject(addIns)) Marshal.ReleaseComObject(addIns);
                }
            }
            return success;
        }

        private bool RestorePaused()
        {
            bool success = true;
            foreach (string id in _pausedBySession.Keys.ToList())
            {
                Office.COMAddIn addIn = null;
                Office.COMAddIns addIns = null;
                try
                {
                    addIns = _application.COMAddIns;
                    addIn = addIns.Item(id);
                    addIn.Connect = true;
                    if (!addIn.Connect) { success = false; continue; }
                    _pausedBySession.Remove(id);
                }
                catch (Exception ex)
                {
                    success = false;
                    ExceptionLogger.Log(ex, "PerformanceMode.Restore." + id);
                }
                finally
                {
                    if (addIn != null && Marshal.IsComObject(addIn)) Marshal.ReleaseComObject(addIn);
                    if (addIns != null && Marshal.IsComObject(addIns)) Marshal.ReleaseComObject(addIns);
                }
            }
            return success;
        }
    }

    internal sealed class ExternalAddInState
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public bool IsConnected { get; set; }
        public bool IsSelected { get; set; }
        public string Status { get; set; }
        public string RestorationCapability { get; set; }
    }
}
