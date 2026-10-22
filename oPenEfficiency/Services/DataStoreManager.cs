using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Xml;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    public static class DataStoreManager
    {
        private const string XML_NAMESPACE = "urn:oPenEfficiency:chartdata";

        public static void SaveData(Presentation presentation, string jsonId, string jsonContent)
        {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrEmpty(jsonId)) throw new ArgumentException("jsonId cannot be null or empty", nameof(jsonId));

            var oldParts = new List<Microsoft.Office.Core.CustomXMLPart>();
            foreach (Microsoft.Office.Core.CustomXMLPart part in presentation.CustomXMLParts.SelectByNamespace(XML_NAMESPACE))
                if (PartMatchesId(part, jsonId)) oldParts.Add(part);

            AtomicStoreTransaction.AddBeforeDelete(
                () => presentation.CustomXMLParts.Add(BuildSafeXml(jsonId, jsonContent)),
                oldParts.ConvertAll(part => (Action)(() => part.Delete())));
        }

        public static string LoadData(Presentation presentation, string jsonId)
        {
            return LoadDataCandidates(presentation, jsonId).FirstOrDefault();
        }

        public static IList<string> LoadDataCandidates(Presentation presentation, string jsonId)
        {
            var candidates = new List<string>();
            if (presentation == null || string.IsNullOrEmpty(jsonId)) return candidates;
            foreach (Microsoft.Office.Core.CustomXMLPart part in presentation.CustomXMLParts.SelectByNamespace(XML_NAMESPACE))
            {
                if (!PartMatchesId(part, jsonId)) continue;
                if (!ExceptionPolicy.TryComProbe(() => part.SelectSingleNode("//*[local-name()='JsonData']"), out Microsoft.Office.Core.CustomXMLNode node) || node == null) continue;
                if (ExceptionPolicy.TryComProbe(() => node.Text, out string text)) candidates.Add(text);
            }
            candidates.Reverse();
            return candidates;
        }

        public static void DeleteData(Presentation presentation, string jsonId)
        {
            if (presentation == null || string.IsNullOrEmpty(jsonId)) return;
            var matches = new List<Microsoft.Office.Core.CustomXMLPart>();
            foreach (Microsoft.Office.Core.CustomXMLPart part in presentation.CustomXMLParts.SelectByNamespace(XML_NAMESPACE))
                if (PartMatchesId(part, jsonId)) matches.Add(part);
            foreach (var part in matches) part.Delete();
        }

        private static bool PartMatchesId(Microsoft.Office.Core.CustomXMLPart part, string jsonId)
        {
            if (part == null) return false;
            string escapedId = SecurityElement.Escape(jsonId);
            return ExceptionPolicy.TryComProbe(() => part.XML, out string xml) &&
                (xml.Contains("id=\"" + escapedId + "\"") || xml.Contains("id='" + escapedId + "'"));
        }

        private static string BuildSafeXml(string jsonId, string jsonContent)
        {
            using (var writerBuffer = new System.IO.StringWriter())
            using (var writer = XmlWriter.Create(writerBuffer, new XmlWriterSettings { OmitXmlDeclaration = true }))
            {
                writer.WriteStartElement("ChartData", XML_NAMESPACE);
                writer.WriteAttributeString("id", jsonId);
                writer.WriteStartElement("JsonData");
                writer.WriteString(jsonContent ?? string.Empty);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.Flush();
                return writerBuffer.ToString();
            }
        }
    }
}
