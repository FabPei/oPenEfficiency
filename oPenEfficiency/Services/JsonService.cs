using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace oPenEfficiency
{
    public sealed class JsonFileStore
    {
        private readonly string _directory;

        public JsonFileStore(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
                throw new ArgumentException("A JSON store directory is required.", nameof(directory));
            _directory = Path.GetFullPath(directory);
        }

        public bool TrySave<T>(T obj, string fileName)
        {
            string tempPath = null;
            try
            {
                string path = ResolvePath(fileName);
                Directory.CreateDirectory(_directory);

                using (var ms = new MemoryStream())
                {
                    var serializer = new DataContractJsonSerializer(typeof(T));
                    serializer.WriteObject(ms, obj);
                    string json = Encoding.UTF8.GetString(ms.ToArray());
                    tempPath = Path.Combine(_directory, Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N") + ".tmp");
                    File.WriteAllText(tempPath, json);
                    if (File.Exists(path)) File.Replace(tempPath, path, null);
                    else File.Move(tempPath, path);
                    tempPath = null;
                }
                return true;
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "JsonFileStore.TrySave." + fileName);
                return false;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempPath))
                {
                    try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { /* Preserve the primary save error if temporary-file cleanup also fails. */ }
                }
            }
        }

        public T Load<T>(string fileName)
        {
            try
            {
                string path = ResolvePath(fileName);
                if (!File.Exists(path)) return default;

                string json = File.ReadAllText(path);
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    var serializer = new DataContractJsonSerializer(typeof(T));
                    return (T)serializer.ReadObject(ms);
                }
            }
            catch (Exception ex)
            {
                oPenEfficiency.Utils.ExceptionLogger.Log(ex, "JsonFileStore.Load." + fileName);
                return default;
            }
        }

        private string ResolvePath(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName) ||
                !string.Equals(fileName, Path.GetFileName(fileName), StringComparison.Ordinal) ||
                fileName.IndexOfAny(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }) >= 0)
            {
                throw new ArgumentException("The JSON file name must not contain a path.", nameof(fileName));
            }
            return Path.Combine(_directory, fileName);
        }
    }

    public static class JsonService
    {
        private const string SidebarConfigFileName = "sidebar_config.json";
        private static readonly string DefaultDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "oPenEfficiency");
        private static readonly JsonFileStore DefaultStore = new JsonFileStore(DefaultDirectory);
        private static readonly SecureSidebarConfigStore DefaultSidebarConfigStore =
            new SecureSidebarConfigStore(
                DefaultStore,
                new DpapiCredentialBlobStore(DefaultDirectory),
                SidebarConfigFileName);

        public static void Save<T>(T obj, string fileName)
        {
            TrySave(obj, fileName);
        }

        public static bool TrySave<T>(T obj, string fileName)
        {
            if (IsSidebarConfig<T>(fileName))
                return DefaultSidebarConfigStore.TrySave((Models.SidebarConfig)(object)obj);

            return DefaultStore.TrySave(obj, fileName);
        }

        public static T Load<T>(string fileName)
        {
            if (IsSidebarConfig<T>(fileName))
                return (T)(object)DefaultSidebarConfigStore.Load();

            return DefaultStore.Load<T>(fileName);
        }

        private static bool IsSidebarConfig<T>(string fileName)
        {
            return typeof(T) == typeof(Models.SidebarConfig) &&
                   string.Equals(fileName, SidebarConfigFileName, StringComparison.OrdinalIgnoreCase);
        }

        public static string Serialize<T>(T obj)
        {
            if (obj == null) return string.Empty;
            
            using (var ms = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                serializer.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static T Deserialize<T>(string json)
        {
            if (string.IsNullOrEmpty(json)) return default(T);
            
            try
            {
                var serializer = new DataContractJsonSerializer(typeof(T));
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    return (T)serializer.ReadObject(ms);
                }
            }
            catch
            {
                return default(T);
            }
        }
    }
}
