using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace oPenEfficiency
{
    public static class JsonService
    {
        private static string GetConfigPath(string fileName)
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "oPenEfficiency", fileName);
        }

        public static void Save<T>(T obj, string fileName)
        {
            try
            {
                string path = GetConfigPath(fileName);
                string directory = Path.GetDirectoryName(path);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                using (var ms = new MemoryStream())
                {
                    var serializer = new DataContractJsonSerializer(typeof(T));
                    serializer.WriteObject(ms, obj);
                    string json = Encoding.UTF8.GetString(ms.ToArray());
                    File.WriteAllText(path, json);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error saving JSON: " + ex.Message);
            }
        }

        public static T Load<T>(string fileName)
        {
            try
            {
                string path = GetConfigPath(fileName);
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
                System.Diagnostics.Debug.WriteLine("Error loading JSON: " + ex.Message);
                return default;
            }
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
