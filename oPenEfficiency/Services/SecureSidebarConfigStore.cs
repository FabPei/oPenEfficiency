using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;

namespace oPenEfficiency
{
    [DataContract]
    internal sealed class SidebarCredentials
    {
        [DataMember(Order = 1)]
        public int Version { get; set; } = 1;

        [DataMember(Order = 2, EmitDefaultValue = false)]
        public string DeepLApiKey { get; set; }

        [DataMember(Order = 3, EmitDefaultValue = false)]
        public string OpenAIApiKey { get; set; }

        [DataMember(Order = 4, EmitDefaultValue = false)]
        public string ClaudeApiKey { get; set; }

        [DataMember(Order = 5, EmitDefaultValue = false)]
        public string LibreTranslateApiKey { get; set; }

        public bool HasAnySecret =>
            !string.IsNullOrEmpty(DeepLApiKey) ||
            !string.IsNullOrEmpty(OpenAIApiKey) ||
            !string.IsNullOrEmpty(ClaudeApiKey) ||
            !string.IsNullOrEmpty(LibreTranslateApiKey);

        public static SidebarCredentials FromConfig(SidebarConfig config)
        {
            return new SidebarCredentials
            {
                DeepLApiKey = EmptyToNull(config?.DeepLApiKey),
                OpenAIApiKey = EmptyToNull(config?.OpenAIApiKey),
                ClaudeApiKey = EmptyToNull(config?.ClaudeApiKey),
                LibreTranslateApiKey = EmptyToNull(config?.LibreTranslateApiKey)
            };
        }

        public void ApplyTo(SidebarConfig config)
        {
            if (config == null) return;
            config.DeepLApiKey = DeepLApiKey ?? string.Empty;
            config.OpenAIApiKey = OpenAIApiKey ?? string.Empty;
            config.ClaudeApiKey = ClaudeApiKey ?? string.Empty;
            config.LibreTranslateApiKey = LibreTranslateApiKey ?? string.Empty;
        }

        public string[] Values()
        {
            return new[] { DeepLApiKey, OpenAIApiKey, ClaudeApiKey, LibreTranslateApiKey };
        }

        private static string EmptyToNull(string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }
    }

    internal enum CredentialReadState
    {
        Missing,
        Success,
        Failed
    }

    internal sealed class CredentialBlobSnapshot
    {
        public bool Captured { get; set; }
        public bool Exists { get; set; }
        public byte[] Bytes { get; set; }
    }

    internal interface ICredentialBlobStore
    {
        CredentialReadState TryRead(out SidebarCredentials credentials);
        CredentialBlobSnapshot Capture();
        bool TryWrite(SidebarCredentials credentials);
        bool TryRestore(CredentialBlobSnapshot snapshot);
    }

    internal sealed class DpapiCredentialBlobStore : ICredentialBlobStore
    {
        internal const string DefaultFileName = "translation_credentials.v1.dat";
        private static readonly byte[] OptionalEntropy = Encoding.UTF8.GetBytes(
            "oPenEfficiency.SidebarCredentials.v1");

        private readonly string _directory;
        private readonly string _path;

        public DpapiCredentialBlobStore(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
                throw new ArgumentException("A credential store directory is required.", nameof(directory));

            _directory = Path.GetFullPath(directory);
            _path = Path.Combine(_directory, DefaultFileName);
        }

        public CredentialReadState TryRead(out SidebarCredentials credentials)
        {
            credentials = null;
            if (!File.Exists(_path)) return CredentialReadState.Missing;

            try
            {
                byte[] protectedBytes = File.ReadAllBytes(_path);
                byte[] plainBytes = ProtectedData.Unprotect(
                    protectedBytes,
                    OptionalEntropy,
                    DataProtectionScope.CurrentUser);

                using (var stream = new MemoryStream(plainBytes))
                {
                    var serializer = new DataContractJsonSerializer(typeof(SidebarCredentials));
                    credentials = (SidebarCredentials)serializer.ReadObject(stream);
                }

                if (credentials == null || credentials.Version != 1)
                    throw new SerializationException("The protected credential format is not supported.");

                return CredentialReadState.Success;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DpapiCredentialBlobStore.TryRead");
                credentials = null;
                return CredentialReadState.Failed;
            }
        }

        public CredentialBlobSnapshot Capture()
        {
            try
            {
                return new CredentialBlobSnapshot
                {
                    Captured = true,
                    Exists = File.Exists(_path),
                    Bytes = File.Exists(_path) ? File.ReadAllBytes(_path) : null
                };
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DpapiCredentialBlobStore.Capture");
                return new CredentialBlobSnapshot { Captured = false };
            }
        }

        public bool TryWrite(SidebarCredentials credentials)
        {
            try
            {
                byte[] plainBytes;
                using (var stream = new MemoryStream())
                {
                    var serializer = new DataContractJsonSerializer(typeof(SidebarCredentials));
                    serializer.WriteObject(stream, credentials ?? new SidebarCredentials());
                    plainBytes = stream.ToArray();
                }

                byte[] protectedBytes = ProtectedData.Protect(
                    plainBytes,
                    OptionalEntropy,
                    DataProtectionScope.CurrentUser);
                return TryWriteBytes(protectedBytes);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DpapiCredentialBlobStore.TryWrite");
                return false;
            }
        }

        public bool TryRestore(CredentialBlobSnapshot snapshot)
        {
            if (snapshot == null || !snapshot.Captured) return false;

            try
            {
                if (snapshot.Exists) return TryWriteBytes(snapshot.Bytes ?? Array.Empty<byte>());
                if (File.Exists(_path)) File.Delete(_path);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DpapiCredentialBlobStore.TryRestore");
                return false;
            }
        }

        private bool TryWriteBytes(byte[] bytes)
        {
            string temporaryPath = null;
            try
            {
                Directory.CreateDirectory(_directory);
                temporaryPath = Path.Combine(
                    _directory,
                    DefaultFileName + "." + Guid.NewGuid().ToString("N") + ".tmp");
                File.WriteAllBytes(temporaryPath, bytes ?? Array.Empty<byte>());

                if (File.Exists(_path)) File.Replace(temporaryPath, _path, null);
                else File.Move(temporaryPath, _path);
                temporaryPath = null;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DpapiCredentialBlobStore.TryWriteBytes");
                return false;
            }
            finally
            {
                if (!string.IsNullOrEmpty(temporaryPath))
                {
                    try
                    {
                        if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
                    }
                    catch
                    {
                        // Preserve the primary storage failure.
                    }
                }
            }
        }
    }

    internal sealed class SecureSidebarConfigStore
    {
        private readonly JsonFileStore _jsonStore;
        private readonly ICredentialBlobStore _credentialStore;
        private readonly string _fileName;

        public SecureSidebarConfigStore(
            JsonFileStore jsonStore,
            ICredentialBlobStore credentialStore,
            string fileName = "sidebar_config.json")
        {
            _jsonStore = jsonStore ?? throw new ArgumentNullException(nameof(jsonStore));
            _credentialStore = credentialStore ?? throw new ArgumentNullException(nameof(credentialStore));
            _fileName = fileName;
        }

        public SidebarConfig Load()
        {
            SidebarConfig config = _jsonStore.Load<SidebarConfig>(_fileName);
            if (config == null)
            {
                return null;
            }

            SidebarCredentials legacyCredentials = SidebarCredentials.FromConfig(config);
            SecretRedactor.RegisterSecrets(legacyCredentials.Values());
            CredentialReadState state = _credentialStore.TryRead(out SidebarCredentials protectedCredentials);

            if (state == CredentialReadState.Success)
            {
                SecretRedactor.RegisterSecrets(protectedCredentials.Values());
                protectedCredentials.ApplyTo(config);
                if (legacyCredentials.HasAnySecret)
                    TryRemoveLegacyCredentials(config);
            }
            else if (legacyCredentials.HasAnySecret)
            {
                // Repair a missing or unreadable blob from the only remaining usable copy.
                if (_credentialStore.TryWrite(legacyCredentials))
                    TryRemoveLegacyCredentials(config);
            }

            SecretRedactor.RegisterSecrets(SidebarCredentials.FromConfig(config).Values());
            return config;
        }

        public bool TrySave(SidebarConfig config)
        {
            if (config == null) return false;

            SidebarCredentials credentials = SidebarCredentials.FromConfig(config);
            SecretRedactor.RegisterSecrets(credentials.Values());
            CredentialBlobSnapshot snapshot = _credentialStore.Capture();
            if (!snapshot.Captured || !_credentialStore.TryWrite(credentials)) return false;

            SidebarConfig sanitized;
            try
            {
                sanitized = CloneWithoutCredentials(config);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SecureSidebarConfigStore.TrySave.Clone");
                RestoreCredentialSnapshot(snapshot);
                return false;
            }

            if (!_jsonStore.TrySave(sanitized, _fileName))
            {
                RestoreCredentialSnapshot(snapshot);
                return false;
            }

            SecretRedactor.RegisterSecrets(credentials.Values());
            return true;
        }

        private void TryRemoveLegacyCredentials(SidebarConfig source)
        {
            try
            {
                SidebarConfig sanitized = CloneWithoutCredentials(source);
                if (_jsonStore.TrySave(sanitized, _fileName)) return;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SecureSidebarConfigStore.Migrate.Clone");
            }

            ExceptionLogger.Log(
                "Protected credentials are available, but legacy plaintext values could not yet be removed from sidebar_config.json.",
                "SecureSidebarConfigStore.Migrate");
        }

        private void RestoreCredentialSnapshot(CredentialBlobSnapshot snapshot)
        {
            if (!_credentialStore.TryRestore(snapshot))
                ExceptionLogger.Log(
                    "Credential rollback failed after the sidebar configuration could not be saved.",
                    "SecureSidebarConfigStore.TrySave");
        }

        private static SidebarConfig CloneWithoutCredentials(SidebarConfig source)
        {
            SidebarConfig clone;
            using (var stream = new MemoryStream())
            {
                var serializer = new DataContractJsonSerializer(typeof(SidebarConfig));
                serializer.WriteObject(stream, source);
                stream.Position = 0;
                clone = (SidebarConfig)serializer.ReadObject(stream);
            }

            clone.DeepLApiKey = null;
            clone.OpenAIApiKey = null;
            clone.ClaudeApiKey = null;
            clone.LibreTranslateApiKey = null;
            return clone;
        }
    }
}
