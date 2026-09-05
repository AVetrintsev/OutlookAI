using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using OutlookAI.Services.CustomActions;

namespace OutlookAI.Services.TextEditing
{
    /// <summary>
    /// Produces the menu fragment returned by an Office Ribbon dynamicMenu
    /// getContent callback. XmlWriter handles attribute escaping; additional
    /// filtering removes characters that XML 1.0 cannot represent.
    /// </summary>
    public static class TextActionMenuBuilder
    {
        public const string RibbonNamespace =
            "http://schemas.microsoft.com/office/2009/07/customui";
        public const string OnActionCallback = "OnTextActionClick";

        public static string Build(IEnumerable<CustomActionDefinition> actions)
        {
            var settings = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                ConformanceLevel = ConformanceLevel.Document,
                Indent = false
            };
            var buffer = new StringBuilder();
            var seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (var writer = XmlWriter.Create(new StringWriter(buffer), settings))
            {
                writer.WriteStartElement("menu", RibbonNamespace);
                var index = 0;
                foreach (var action in actions ?? new CustomActionDefinition[0])
                {
                    if (!CanRender(action) || !seenIds.Add(action.Id)) continue;

                    writer.WriteStartElement("button", RibbonNamespace);
                    writer.WriteAttributeString("id", "outlookAiTextAction" + index++);
                    writer.WriteAttributeString("label", Limit(
                        SanitizeXmlValue(action.Title.Trim()), 255));
                    writer.WriteAttributeString("tag", action.Id);
                    writer.WriteAttributeString("onAction", OnActionCallback);
                    if (!string.IsNullOrWhiteSpace(action.Description))
                    {
                        writer.WriteAttributeString("supertip", Limit(
                            SanitizeXmlValue(action.Description.Trim()), 1024));
                    }
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            return buffer.ToString();
        }

        public static string BuildMenuXml(IEnumerable<CustomActionDefinition> actions)
        {
            return Build(actions);
        }

        private static bool CanRender(CustomActionDefinition action)
        {
            return TextActionCatalog.IsTextAction(action)
                && !action.Disabled
                && !string.IsNullOrWhiteSpace(action.Id)
                && !string.IsNullOrWhiteSpace(action.Title)
                && string.Equals(
                    action.Id,
                    SanitizeXmlValue(action.Id),
                    StringComparison.Ordinal);
        }

        private static string SanitizeXmlValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            var result = new StringBuilder(value.Length);
            for (var i = 0; i < value.Length; i++)
            {
                var ch = value[i];
                if (char.IsHighSurrogate(ch)
                    && i + 1 < value.Length
                    && char.IsLowSurrogate(value[i + 1]))
                {
                    result.Append(ch);
                    result.Append(value[++i]);
                    continue;
                }
                if (ch == '\t'
                    || ch == '\n'
                    || ch == '\r'
                    || (ch >= 0x20 && ch <= 0xD7FF)
                    || (ch >= 0xE000 && ch <= 0xFFFD))
                {
                    result.Append(ch);
                }
            }
            return result.ToString();
        }

        private static string Limit(string value, int maxCharacters)
        {
            value = value ?? "";
            if (value.Length <= maxCharacters) return value;
            var length = maxCharacters;
            if (length > 0
                && char.IsHighSurrogate(value[length - 1])
                && length < value.Length
                && char.IsLowSurrogate(value[length]))
            {
                length--;
            }
            return value.Substring(0, length);
        }
    }
}
