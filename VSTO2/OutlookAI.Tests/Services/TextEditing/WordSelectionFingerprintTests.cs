using OutlookAI.Services.TextEditing;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    public sealed class WordSelectionFingerprintTests
    {
        private const string Xml = "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>"
            + "<w:body><w:p w:rsidRDefault='00000001'><w:r><w:rPr><w:b/></w:rPr><w:t>Text</w:t></w:r></w:p></w:body></w:document>";

        [Fact]
        public void IgnoresOnlyWordRevisionSessionMetadata()
        {
            Assert.Equal(WordSelectionEditor.NormalizeFingerprintXml(Xml),
                WordSelectionEditor.NormalizeFingerprintXml(Xml.Replace("00000001", "12345678")));
        }

        [Theory]
        [InlineData("Text", "Changed")]
        [InlineData("<w:b/>", "<w:i/>")]
        public void PreservesTextAndFormattingChanges(string before, string after)
        {
            Assert.NotEqual(WordSelectionEditor.NormalizeFingerprintXml(Xml),
                WordSelectionEditor.NormalizeFingerprintXml(Xml.Replace(before, after)));
        }

        [Fact]
        public void PreservesHyperlinkTargetsAndOtherNamespaces()
        {
            const string links = "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
                + "<Relationship Id='rId1' Target='https://example.org/first'/></Relationships>";
            Assert.NotEqual(WordSelectionEditor.NormalizeFingerprintXml(links),
                WordSelectionEditor.NormalizeFingerprintXml(links.Replace("/first", "/second")));
            const string unrelated = "<item xmlns:x='urn:example' x:rsid='1'/>";
            Assert.NotEqual(WordSelectionEditor.NormalizeFingerprintXml(unrelated),
                WordSelectionEditor.NormalizeFingerprintXml(unrelated.Replace("'1'", "'2'")));
        }

        [Fact]
        public void IgnoresGeneratedDocumentIdentityButNotDocumentContent()
        {
            const string first = "<settings xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'><w15:docId w15:val='one'/><text>same</text></settings>";
            Assert.Equal(WordSelectionEditor.NormalizeFingerprintXml(first),
                WordSelectionEditor.NormalizeFingerprintXml(first.Replace("'one'", "'two'")));
            Assert.NotEqual(WordSelectionEditor.NormalizeFingerprintXml(first),
                WordSelectionEditor.NormalizeFingerprintXml(first.Replace(">same<", ">changed<")));
        }
    }
}
