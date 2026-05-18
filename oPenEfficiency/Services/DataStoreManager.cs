using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using System.Xml;

namespace oPenEfficiency.Services
{
    public static class DataStoreManager
    {
        private const string XML_NAMESPACE = "urn:oPenEfficiency:chartdata";

        public static void SaveData(Presentation presentation, string jsonId, string jsonContent)
        {
            if (string.IsNullOrEmpty(jsonId))
                throw new ArgumentException("jsonId cannot be null or empty", nameof(jsonId));

            // Escape jsonId for safe use in XPath and XML
            string safeJsonId = EscapeForXmlAttribute(jsonId);

            var parts = presentation.CustomXMLParts.SelectByNamespace(XML_NAMESPACE);
            foreach (Microsoft.Office.Core.CustomXMLPart part in parts)
            {
                // Use safer check that doesn't rely on string interpolation
                if (part.XML.Contains("id=\"" + safeJsonId + "\"") ||
                    part.XML.Contains("id='" + safeJsonId + "'"))
                {
                    part.Delete();
                }
            }

            // Build XML safely using XmlWriter to prevent injection
            string xmlString = BuildSafeXml(safeJsonId, jsonContent);
            presentation.CustomXMLParts.Add(xmlString);
        }

        public static string LoadData(Presentation presentation, string jsonId)
        {
            if (string.IsNullOrEmpty(jsonId))
                return null;

            // Escape jsonId for safe comparison
            string safeJsonId = EscapeForXmlAttribute(jsonId);

            var parts = presentation.CustomXMLParts.SelectByNamespace(XML_NAMESPACE);
            foreach (Microsoft.Office.Core.CustomXMLPart part in parts)
            {
                // Use safer check that doesn't rely on string interpolation
                if (part.XML.Contains("id=\"" + safeJsonId + "\"") ||
                    part.XML.Contains("id='" + safeJsonId + "'"))
                {
                    var node = part.SelectSingleNode("//*[local-name()='JsonData']");
                    if (node != null)
                    {
                        return node.Text;
                    }
                }
            }
            return null;
        }

        private static string EscapeForXmlAttribute(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            // Replace XML special characters in attribute values
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        private static string BuildSafeXml(string escapedJsonId, string jsonContent)
        {
            // Handle CDATA: if content contains ]]>, we need to split it
            // CDATA sections cannot contain the sequence ]]>
            string safeContent = jsonContent ?? string.Empty;

            using (var stringWriter = new System.IO.StringWriter())
            using (var xmlWriter = XmlWriter.Create(stringWriter, new XmlWriterSettings { OmitXmlDeclaration = true }))
            {
                xmlWriter.WriteStartElement("ChartData", XML_NAMESPACE);
                xmlWriter.WriteAttributeString("id", escapedJsonId);
                xmlWriter.WriteStartElement("JsonData");

                // Write content as CDATA, handling the ]]> edge case
                if (safeContent.Contains("]]>"))
                {
                    // Split content around ]]> and write as separate CDATA sections
                    var parts = safeContent.Split(new[] { "]]" }, StringSplitOptions.None);
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (i > 0)
                            xmlWriter.WriteRaw("]]");
                        if (i < parts.Length - 1 || safeContent.EndsWith("]]"))
                            xmlWriter.WriteRaw(parts[i] + ">");
                        else
                            xmlWriter.WriteRaw(parts[i]);
                    }
                }
                else
                {
                    xmlWriter.WriteCData(safeContent);
                }

                xmlWriter.WriteEndElement(); // JsonData
                xmlWriter.WriteEndElement(); // ChartData
                xmlWriter.Flush();

                return stringWriter.ToString();
            }
        }
    }
}
