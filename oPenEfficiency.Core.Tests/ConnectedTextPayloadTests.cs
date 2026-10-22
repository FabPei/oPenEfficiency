using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using oPenEfficiency.OneNote;

namespace oPenEfficiency.Core.Tests
{
    /// <summary>
    /// QA-S9 external-data-wave coverage for the connected text label payload
    /// (QA9-BtnConnectedTextLabel-002): the OneNote HTML marker contract, the
    /// lastModified concurrency token read and the persisted target-record JSON shape,
    /// all documented compatibility contracts of FEAT-007 ON-S0 used by
    /// ConnectedTextLabelFeature and the OneNote services.
    /// </summary>
    [TestClass]
    public sealed class ConnectedTextPayloadTests
    {
        [TestMethod]
        public void BuildHtmlEncodesTextAndPreservesLineBreaks()
        {
            string html = ConnectedTextPayload.BuildHtml("src-1", "<b>Bold</b>\r\nsecond\nthird");
            Assert.AreEqual(
                "<span data-ope-source=\"src-1\">&lt;b&gt;Bold&lt;/b&gt;<br />second<br />third</span>",
                html);
        }

        [TestMethod]
        public void BuildHtmlToleratesNullText()
        {
            Assert.AreEqual("<span data-ope-source=\"src-1\"></span>",
                ConnectedTextPayload.BuildHtml("src-1", null));
        }

        [TestMethod]
        public void SourceMarkerIsTheExactLegacySubstring()
        {
            Assert.AreEqual("data-ope-source=\"src-1\"", ConnectedTextPayload.SourceMarker("src-1"));
            StringAssert.Contains(ConnectedTextPayload.BuildHtml("src-1", "x"), ConnectedTextPayload.SourceMarker("src-1"));
        }

        [TestMethod]
        public void ReadLastModifiedRoundTripsTheRoundtripKindFormat()
        {
            var page = XDocument.Parse(
                "<page lastModifiedTime=\"2026-09-18T10:30:00.000Z\"><title>t</title></page>");
            DateTime value = ConnectedTextPayload.ReadLastModified(page);
            Assert.AreEqual(new DateTime(2026, 9, 18, 10, 30, 0, DateTimeKind.Utc), value);
        }

        [TestMethod]
        public void ReadLastModifiedFailsClosedToMinValue()
        {
            Assert.AreEqual(DateTime.MinValue, ConnectedTextPayload.ReadLastModified(null));
            Assert.AreEqual(DateTime.MinValue, ConnectedTextPayload.ReadLastModified(XDocument.Parse("<page><title>t</title></page>")));
            Assert.AreEqual(DateTime.MinValue, ConnectedTextPayload.ReadLastModified(XDocument.Parse("<page lastModifiedTime=\"not-a-date\"/>")));
        }

        [TestMethod]
        public void TargetRecordKeepsTheLegacyJsonPropertyNames()
        {
            var record = new ConnectedTextTargetRecord
            {
                SourceId = "src-1",
                Host = "OneNote",
                DocumentPath = "C:\\Notebooks\\n.one",
                TargetId = "page-id-9"
            };
            string json = Serialize(record);
            StringAssert.Contains(json, "SourceId");
            StringAssert.Contains(json, "Host");
            StringAssert.Contains(json, "DocumentPath");
            StringAssert.Contains(json, "TargetId");

            ConnectedTextTargetRecord roundTrip = Deserialize(json);
            Assert.AreEqual("src-1", roundTrip.SourceId);
            Assert.AreEqual("OneNote", roundTrip.Host);
            Assert.AreEqual("C:\\Notebooks\\n.one", roundTrip.DocumentPath);
            Assert.AreEqual("page-id-9", roundTrip.TargetId);
        }

        private static string Serialize(ConnectedTextTargetRecord record)
        {
            var serializer = new DataContractJsonSerializer(typeof(ConnectedTextTargetRecord));
            using (var stream = new MemoryStream())
            {
                serializer.WriteObject(stream, record);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }

        private static ConnectedTextTargetRecord Deserialize(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(ConnectedTextTargetRecord));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (ConnectedTextTargetRecord)serializer.ReadObject(stream);
            }
        }
    }
}
