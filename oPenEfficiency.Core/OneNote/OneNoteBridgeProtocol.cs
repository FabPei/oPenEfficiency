using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;

namespace oPenEfficiency.OneNote
{
    /// <summary>
    /// Typed, versioned contract for the local OneNote host. The protocol deliberately
    /// has no arbitrary COM invocation or arbitrary XML update operation.
    /// </summary>
    public static class OneNoteBridgeProtocol
    {
        public const int Version = 1;
        public const int MaxMessageBytes = 1024 * 1024;
        public const string GetActivePage = "GetActivePage";
        public const string ReadPage = "ReadPage";
        public const string AppendConnectedText = "AppendConnectedText";
        public const string RefreshConnectedTexts = "RefreshConnectedTexts";

        public static string GetPipeNameForCurrentUser()
        {
            string sid = WindowsIdentity.GetCurrent()?.User?.Value;
            if (string.IsNullOrWhiteSpace(sid))
                throw new InvalidOperationException("The current Windows user SID is unavailable.");

            using (SHA256 sha = SHA256.Create())
            {
                byte[] digest = sha.ComputeHash(Encoding.UTF8.GetBytes(sid));
                var suffix = new StringBuilder(24);
                for (int index = 0; index < 12; index++) suffix.Append(digest[index].ToString("x2"));
                return "oPenEfficiency.OneNote." + suffix;
            }
        }

        public static void WriteRequest(Stream stream, OneNoteBridgeRequest request) => Write(stream, request);

        public static OneNoteBridgeRequest ReadRequest(Stream stream) => Read<OneNoteBridgeRequest>(stream);

        public static void WriteResponse(Stream stream, OneNoteBridgeResponse response) => Write(stream, response);

        public static OneNoteBridgeResponse ReadResponse(Stream stream) => Read<OneNoteBridgeResponse>(stream);

        private static void Write<T>(Stream stream, T message)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (message == null) throw new ArgumentNullException(nameof(message));

            byte[] payload;
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var buffer = new MemoryStream())
            {
                serializer.WriteObject(buffer, message);
                payload = buffer.ToArray();
            }

            if (payload.Length <= 0 || payload.Length > MaxMessageBytes)
                throw new InvalidDataException("The OneNote bridge message has an invalid size.");

            byte[] length = BitConverter.GetBytes(payload.Length);
            stream.Write(length, 0, length.Length);
            stream.Write(payload, 0, payload.Length);
            stream.Flush();
        }

        private static T Read<T>(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            byte[] lengthBytes = ReadExactly(stream, sizeof(int));
            int length = BitConverter.ToInt32(lengthBytes, 0);
            if (length <= 0 || length > MaxMessageBytes)
                throw new InvalidDataException("The OneNote bridge message has an invalid size.");

            byte[] payload = ReadExactly(stream, length);
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var buffer = new MemoryStream(payload, writable: false))
            {
                object value = serializer.ReadObject(buffer);
                if (!(value is T typed)) throw new SerializationException("The OneNote bridge message is invalid.");
                return typed;
            }
        }

        private static byte[] ReadExactly(Stream stream, int length)
        {
            var buffer = new byte[length];
            int offset = 0;
            while (offset < length)
            {
                int read = stream.Read(buffer, offset, length - offset);
                if (read <= 0) throw new EndOfStreamException("The OneNote bridge connection closed unexpectedly.");
                offset += read;
            }
            return buffer;
        }
    }

    [DataContract]
    public sealed class OneNoteBridgeRequest
    {
        [DataMember(Order = 1)] public int Version { get; set; } = OneNoteBridgeProtocol.Version;
        [DataMember(Order = 2)] public string Operation { get; set; }
        [DataMember(Order = 3, EmitDefaultValue = false)] public string PageId { get; set; }
        [DataMember(Order = 4, EmitDefaultValue = false)] public string SourceId { get; set; }
        [DataMember(Order = 5, EmitDefaultValue = false)] public string Text { get; set; }
        [DataMember(Order = 6, EmitDefaultValue = false)] public List<OneNoteBridgeTarget> Targets { get; set; }
    }

    [DataContract]
    public sealed class OneNoteBridgeTarget
    {
        [DataMember(Order = 1)] public string SourceId { get; set; }
        [DataMember(Order = 2, EmitDefaultValue = false)] public string TargetId { get; set; }
        [DataMember(Order = 3)] public string Text { get; set; }
    }

    [DataContract]
    public sealed class OneNoteBridgeResponse
    {
        [DataMember(Order = 1)] public int Version { get; set; } = OneNoteBridgeProtocol.Version;
        [DataMember(Order = 2)] public bool Success { get; set; }
        [DataMember(Order = 3, EmitDefaultValue = false)] public string Error { get; set; }
        [DataMember(Order = 4, EmitDefaultValue = false)] public string PageId { get; set; }
        [DataMember(Order = 5, EmitDefaultValue = false)] public string PageXml { get; set; }
        [DataMember(Order = 6, EmitDefaultValue = false)] public string LastModifiedTime { get; set; }
        [DataMember(Order = 7, EmitDefaultValue = false)] public int ReplacedCount { get; set; }
    }
}
