using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OutlookAI.Services.TextEditing
{
    /// <summary>All calls must run on Outlook's STA. Owns only duplicate ranges and a font snapshot.</summary>
    internal sealed class WordSelectionEditor : IDisposable
    {
        private readonly object _document;
        private readonly object _selection;
        private readonly List<object> _segments = new List<object>();
        private object _font;
        private string _beforeFingerprint;
        private string _afterDocumentFingerprint;
        private int _originalSpan;
        private int _appliedSpan;

        private WordSelectionEditor(object document, object selection)
        {
            _document = document;
            _selection = selection;
        }

        public ProtectedTextPlan Plan { get; private set; }

        public static WordSelectionEditor Capture(object documentObject, object rangeObject)
        {
            var editor = new WordSelectionEditor(documentObject, rangeObject);
            try
            {
                dynamic document = documentObject;
                dynamic range = rangeObject;
                // Document.Range coordinates refer to the main story, never a text box/header.
                if ((int)range.StoryType != 1) throw Unsupported(); // wdMainTextStory
                if ((bool)document.TrackRevisions)
                    throw new InvalidOperationException("Сначала отключите запись исправлений Word для этого документа.");
                var start = (int)range.Start;
                var end = (int)range.End;
                editor._originalSpan = end - start;
                RejectObjects(documentObject, rangeObject, start, end);
                var anchors = new List<Tuple<int, int>>();
                object fieldsObject = document.Fields;
                try
                {
                    dynamic fields = fieldsObject;
                    for (var i = 1; i <= (int)fields.Count; i++)
                    {
                        object fieldObject = fields[i];
                        object codeObject = null;
                        object resultObject = null;
                        try
                        {
                            dynamic field = fieldObject;
                            codeObject = field.Code;
                            resultObject = field.Result;
                            var fieldStart = (int)((dynamic)codeObject).Start - 1;
                            var fieldEnd = (int)((dynamic)resultObject).End + 1;
                            if (fieldEnd <= start || fieldStart >= end) continue;
                            if ((int)field.Type != 88) throw Unsupported(); // wdFieldHyperlink
                            anchors.Add(Tuple.Create(Math.Max(start, fieldStart), Math.Min(end, fieldEnd)));
                        }
                        finally { Release(resultObject); Release(codeObject); Release(fieldObject); }
                    }
                }
                finally { Release(fieldsObject); }

                var texts = new List<string>();
                var labels = new List<string>();
                var cursor = start;
                foreach (var anchor in anchors.OrderBy(pair => pair.Item1))
                {
                    if (anchor.Item1 < cursor) throw Unsupported();
                    editor.AddSegment(cursor, anchor.Item1, texts);
                    object linkRange = document.Range(anchor.Item1, anchor.Item2);
                    try { labels.Add((string)((dynamic)linkRange).Text ?? ""); }
                    finally { Release(linkRange); }
                    cursor = anchor.Item2;
                }
                editor.AddSegment(cursor, end, texts);
                editor.Plan = new ProtectedTextPlan(texts, labels);
                if (editor.Plan.OriginalText != ((string)range.Text ?? "")) throw Unsupported();
                // Font.Duplicate is an independent snapshot, not a live range font.
                var fontStart = start;
                foreach (var segment in editor._segments)
                {
                    if ((int)((dynamic)segment).End > (int)((dynamic)segment).Start)
                    {
                        fontStart = (int)((dynamic)segment).Start;
                        break;
                    }
                }
                object firstChar = document.Range(fontStart, Math.Min(fontStart + 1, end));
                object fontObject = null;
                try
                {
                    fontObject = ((dynamic)firstChar).Font;
                    editor._font = ((dynamic)fontObject).Duplicate;
                }
                finally { Release(fontObject); Release(firstChar); }
                editor._beforeFingerprint = Fingerprint(rangeObject);
                return editor;
            }
            catch { editor.Dispose(); throw; }
        }

        public bool IsCurrent => Fingerprint(_selection) == _beforeFingerprint;

        public string Apply(string modelReplacement, string name)
        {
            var parts = Plan.ParseReplacement(modelReplacement);
            var expected = Plan.GetText(parts);
            var changes = Plan.GetChanges(parts); // Validate all plans before the first write.
            dynamic document = _document;
            dynamic selection = _selection;
            var start = (int)selection.Start;
            var end = (int)selection.End + expected.Length - Plan.OriginalText.Length;
            object applicationObject = document.Application;
            object undoObject = null;
            var started = false;
            var changed = false;
            try
            {
                undoObject = ((dynamic)applicationObject).UndoRecord;
                dynamic undo = undoObject;
                if ((int)undo.CustomRecordLevel != 0)
                    throw new InvalidOperationException("Word уже выполняет другую групповую правку. Повторите позже.");
                undo.StartCustomRecord((name ?? "OutlookAI").Substring(0, Math.Min(64, (name ?? "OutlookAI").Length)));
                started = true;
                for (var i = changes.Length - 1; i >= 0; i--)
                {
                    var segmentStart = (int)((dynamic)_segments[i]).Start;
                    foreach (var change in changes[i].Reverse())
                    {
                        object patchObject = document.Range(segmentStart + change.Start, segmentStart + change.Start + change.Length);
                        try
                        {
                            dynamic patch = patchObject;
                            patch.Text = change.Text;
                            changed = true;
                            if (change.Text.Length > 0)
                            {
                                patch.SetRange(segmentStart + change.Start, segmentStart + change.Start + change.Text.Length);
                                patch.Font = _font;
                            }
                        }
                        finally { Release(patchObject); }
                    }
                }
                selection.SetRange(start, end);
                if (((string)selection.Text ?? "") != expected)
                    throw new InvalidOperationException("Word изменил текст при вставке. Замена отменена.");
                undo.EndCustomRecord();
                started = false;
                _appliedSpan = end - start;
                _afterDocumentFingerprint = DocumentFingerprint();
                selection.Select();
                return expected;
            }
            catch
            {
                if (started)
                {
                    ((dynamic)undoObject).EndCustomRecord();
                    started = false;
                }
                // This group contains only our writes, and the STA was not yielded.
                if (changed)
                {
                    document.Undo(1);
                    selection.SetRange(start, start + _originalSpan);
                    selection.Select();
                }
                throw;
            }
            finally { Release(undoObject); Release(applicationObject); }
        }

        public bool Undo()
        {
            dynamic range = _selection;
            if (((string)range.Text ?? "") == Plan.OriginalText) return true;
            // Do not undo someone else's later edit, even if it is outside our selection.
            if (string.IsNullOrEmpty(_afterDocumentFingerprint)
                || DocumentFingerprint() != _afterDocumentFingerprint) return false;
            var start = (int)range.Start;
            if ((int)range.End - start != _appliedSpan) return false;
            if (!((bool)((dynamic)_document).Undo(1))) return false;
            range.SetRange(start, start + _originalSpan);
            range.Select();
            return ((string)range.Text ?? "") == Plan.OriginalText;
        }

        public void Focus()
        {
            ((dynamic)_document).Activate();
            ((dynamic)_selection).Select();
        }

        private void AddSegment(int start, int end, List<string> texts)
        {
            object segment = ((dynamic)_document).Range(start, end);
            _segments.Add(segment);
            var text = (string)((dynamic)segment).Text ?? "";
            if (text.Length != end - start || text.IndexOfAny(new[] { '\a', '\u0001', '\u0002', '\u0013', '\u0014', '\u0015' }) >= 0)
                throw Unsupported();
            texts.Add(text);
        }

        private static void RejectObjects(object documentObject, object rangeObject, int start, int end)
        {
            dynamic range = rangeObject;
            object inlineShapes = range.InlineShapes;
            object controls = range.ContentControls;
            object shapesObject = ((dynamic)documentObject).Shapes;
            try
            {
                if ((int)((dynamic)inlineShapes).Count > 0 || (int)((dynamic)controls).Count > 0) throw Unsupported();
                dynamic shapes = shapesObject;
                for (var i = 1; i <= (int)shapes.Count; i++)
                {
                    object shape = shapes[i];
                    object anchor = null;
                    try
                    {
                        anchor = ((dynamic)shape).Anchor;
                        var position = (int)((dynamic)anchor).Start;
                        if (position >= start && position < end) throw Unsupported();
                    }
                    finally { Release(anchor); Release(shape); }
                }
            }
            finally { Release(shapesObject); Release(controls); Release(inlineShapes); }
        }

        private string DocumentFingerprint()
        {
            object content = ((dynamic)_document).Content;
            try { return Fingerprint(content); }
            finally { Release(content); }
        }

        private static string Fingerprint(object range)
        {
            using (var hash = SHA256.Create())
                return Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(
                    NormalizeFingerprintXml((string)((dynamic)range).WordOpenXML))));
        }

        internal static string NormalizeFingerprintXml(string xml)
        {
            // Range.WordOpenXML synthesizes new revision session IDs on every read,
            // even for an unchanged range. They are bookkeeping, not document edits.
            // Keep all text, formatting, relationships, fields and objects in the hash.
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            foreach (var attribute in document.Descendants().Attributes()
                .Where(value => value.Name.Namespace == word
                    && value.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).ToArray())
                attribute.Remove();
            foreach (var sessions in document.Descendants(word + "rsids").ToArray()) sessions.Remove();
            // The temporary package also receives a fresh document ID on export.
            foreach (var identity in document.Descendants().Where(value => value.Name.LocalName == "docId"
                && (value.Name.NamespaceName == "http://schemas.microsoft.com/office/word/2010/wordml"
                    || value.Name.NamespaceName == "http://schemas.microsoft.com/office/word/2012/wordml")).ToArray())
                identity.Remove();
            return document.ToString(SaveOptions.DisableFormatting);
        }

        private static InvalidOperationException Unsupported() => new InvalidOperationException(
            "Выделите обычный текст или текст внутри одной ячейки, без объектов и служебных полей. Гиперссылки поддерживаются.");

        public void Dispose()
        {
            foreach (var segment in _segments) Release(segment);
            _segments.Clear();
            Release(_font); _font = null;
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
