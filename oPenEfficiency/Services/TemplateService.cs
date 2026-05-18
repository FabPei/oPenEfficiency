using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Models;

namespace oPenEfficiency.Services
{
    public class TemplateService
    {
        private const string ConfigFileName = "templates_config.json";
        private TemplateConfig _config;

        public TemplateService()
        {
            _config = JsonService.Load<TemplateConfig>(ConfigFileName) ?? new TemplateConfig();
            if (_config.BlacklistedPaths == null) _config.BlacklistedPaths = new List<string>();
        }

        public void SaveConfig()
        {
            JsonService.Save(_config, ConfigFileName);
        }

        public List<TemplateItem> GetAllTemplates()
        {
            var templates = new List<TemplateItem>();

            // 1. Add custom templates from config
            if (_config.CustomTemplates != null)
                templates.AddRange(_config.CustomTemplates);

            // 2. Add templates from watched directories
            if (_config.WatchedDirectories != null)
            {
                foreach (var dir in _config.WatchedDirectories)
                {
                    if (Directory.Exists(dir))
                    {
                        templates.AddRange(GetTemplatesInDirectory(dir));
                    }
                }
            }

            // 3. Add default PowerPoint template locations (if not hidden)
            if (!_config.HideDefaultLocations)
            {
                foreach (var dir in GetDefaultTemplateLocations())
                {
                    if (Directory.Exists(dir))
                    {
                        templates.AddRange(GetTemplatesInDirectory(dir));
                    }
                }
            }

            // De-duplicate by path and filter blacklist
            var blacklist = new HashSet<string>(_config.BlacklistedPaths.Select(p => p.ToLower()));
            return templates
                .Where(t => !blacklist.Contains(t.Path.ToLower()))
                .GroupBy(t => t.Path.ToLower())
                .Select(g => g.First())
                .ToList();
        }

        private List<TemplateItem> GetTemplatesInDirectory(string dir)
        {
            var items = new List<TemplateItem>();
            try
            {
                string[] extensions = { "*.potx", "*.pot", "*.pptx", "*.pptm" };
                foreach (var ext in extensions)
                {
                    foreach (var file in Directory.GetFiles(dir, ext, SearchOption.TopDirectoryOnly))
                    {
                        items.Add(new TemplateItem { Name = Path.GetFileNameWithoutExtension(file), Path = file });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error scanning directory {dir}: {ex.Message}");
            }
            return items;
        }

        private List<string> GetDefaultTemplateLocations()
        {
            var paths = new List<string>();
            paths.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Templates"));
            paths.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Custom Office Templates"));
            return paths;
        }

        public void AddCustomTemplate(string path)
        {
            if (File.Exists(path))
            {
                if (_config.CustomTemplates == null) _config.CustomTemplates = new List<TemplateItem>();
                if (!_config.CustomTemplates.Any(t => t.Path.Equals(path, StringComparison.OrdinalIgnoreCase)))
                {
                    _config.CustomTemplates.Add(new TemplateItem { Name = Path.GetFileNameWithoutExtension(path), Path = path });
                    SaveConfig();
                }
            }
        }

        public void RemoveCustomTemplate(string path)
        {
            if (_config.CustomTemplates != null)
            {
                _config.CustomTemplates.RemoveAll(t => t.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
                SaveConfig();
            }
        }

        public void AddToBlacklist(string path)
        {
            if (!_config.BlacklistedPaths.Contains(path, StringComparer.OrdinalIgnoreCase))
            {
                _config.BlacklistedPaths.Add(path);
                SaveConfig();
            }
        }

        public void ClearBlacklist()
        {
            _config.BlacklistedPaths.Clear();
            SaveConfig();
        }

        public void OpenTemplate(PowerPointManager manager, string path)
        {
            try
            {
                var app = manager.GetApplication();
                app.Presentations.Open(path, Untitled: Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Failed to open template: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public byte[] GetThumbnail(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                
                // Legacy .pot files don't support ZIP peek
                if (path.EndsWith(".pot", StringComparison.OrdinalIgnoreCase)) return null;

                using (ZipArchive archive = ZipFile.OpenRead(path))
                {
                    var entry = archive.GetEntry("docProps/thumbnail.jpeg");
                    if (entry == null) return null;

                    using (Stream stream = entry.Open())
                    using (MemoryStream ms = new MemoryStream())
                    {
                        stream.CopyTo(ms);
                        return ms.ToArray();
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public void SetHideDefaultLocations(bool hide)
        {
            _config.HideDefaultLocations = hide;
            SaveConfig();
        }

        public void AddWatchedDirectory(string path)
        {
            if (Directory.Exists(path))
            {
                if (_config.WatchedDirectories == null) _config.WatchedDirectories = new List<string>();
                if (!_config.WatchedDirectories.Contains(path))
                {
                    _config.WatchedDirectories.Add(path);
                    SaveConfig();
                }
            }
        }

        public void RemoveWatchedDirectory(string path)
        {
            if (_config.WatchedDirectories != null)
            {
                _config.WatchedDirectories.Remove(path);
                SaveConfig();
            }
        }

        public TemplateConfig GetConfig() => _config;
    }
}
