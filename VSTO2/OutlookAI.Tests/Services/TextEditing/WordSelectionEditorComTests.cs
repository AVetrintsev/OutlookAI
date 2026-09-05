using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using OutlookAI.Services.TextEditing;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    // Explicit opt-in: these tests start a private Word instance, never attach to a running one.
    // Run on an interactive Windows desktop with Word installed and first-run dialogs completed.
    public sealed class WordComFactAttribute : FactAttribute
    {
        public WordComFactAttribute()
        {
            if (Environment.GetEnvironmentVariable("OUTLOOKAI_RUN_WORD_COM_TESTS") != "1")
                Skip = "Set OUTLOOKAI_RUN_WORD_COM_TESTS=1 to run isolated, unsaved Word COM tests.";
        }
    }

    [CollectionDefinition("Word COM", DisableParallelization = true)]
    public sealed class WordComCollection { }

    [Collection("Word COM")]
    public sealed class WordSelectionEditorComTests
    {
        private const int AccentColor = 0x563412; // Word wdColor is an RGB value, not a theme index.

        [WordComFact]
        public void ConsecutiveReadsOfUnchangedRangeHaveTheSameFingerprint()
        {
            RunInSta(word =>
            {
                word.SetText("Unchanged sample");
                dynamic range = word.BodyWithoutFinalParagraph();
                string first = range.WordOpenXML;
                string second = range.WordOpenXML;
                Assert.Equal(WordSelectionEditor.NormalizeFingerprintXml(first), WordSelectionEditor.NormalizeFingerprintXml(second));
            });
        }

        [WordComFact]
        public void ReplacementPreservesTwoHyperlinksFormattingAndSelectionAndIsOneNativeUndo()
        {
            RunInSta(word =>
            {
                const string original = "Before First link keep bold color Second link After";
                const string expected = "Please read First link keep bold color Second link today";
                word.SetText(original);
                word.AddLink("Second link", "https://example.org/second?q=2", "section-two");
                word.AddLink("First link", "https://example.org/first?q=1", "section-one");
                word.SetEmphasis("bold color", AccentColor);
                object range = word.BodyWithoutFinalParagraph();
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    string replacement = editor.Plan.ModelText.Replace("Before", "Please read").Replace("After", "today");
                    Assert.Equal(expected, editor.Apply(replacement, "OutlookAI COM test"));
                    Assert.Equal(expected + "\r", word.Text);
                    word.AssertLinks();
                    word.AssertEmphasis("bold color", AccentColor);
                    word.AssertSelectionMatches(range, expected);

                    // This is Word's native Undo, the operation used by Ctrl+Z, not editor.Undo().
                    Assert.True((bool)((dynamic)word.Document).Undo(1));
                    Assert.Equal(original + "\r", word.Text);
                    word.AssertLinks();
                    word.AssertEmphasis("bold color", AccentColor);
                }
            });
        }

        [WordComFact]
        public void SelectionStartingInsideHyperlinkPreservesItsFullCaptionAndAddress()
        {
            RunInSta(word =>
            {
                word.SetText("Prefix Alpha Link old suffix");
                word.AddLink("Alpha Link", "https://example.org/partial?q=1", "anchor");
                dynamic caption = word.Find("Alpha Link");
                dynamic tail = word.Find("old");
                object range = word.Range((int)caption.Start + "Alpha ".Length, (int)tail.End);
                ((dynamic)range).Select();
                // Word can expand a selection across a hidden field boundary. The controller
                // captures the actual UI Selection.Range, not the synthetic requested Range.
                object nativeRange = word.Own(((dynamic)word.Selection()).Range);
                range = word.Own(((dynamic)nativeRange).Duplicate);
                var selectedBefore = (string)((dynamic)range).Text;
                Assert.EndsWith("Link old", selectedBefore);
                var expectedSelection = selectedBefore.Replace("old", "revised");
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    Assert.Equal(expectedSelection, editor.Apply(editor.Plan.ModelText.Replace("old", "revised"), "Partial link"));
                    word.AssertSelectionMatches(range, expectedSelection);
                    Assert.Equal("Prefix Alpha Link revised suffix\r", word.Text);
                    word.AssertLink(1, "Alpha Link", "https://example.org/partial?q=1", "anchor");
                    Assert.True((bool)((dynamic)word.Document).Undo(1));
                    Assert.Equal("Prefix Alpha Link old suffix\r", word.Text);
                    word.AssertLink(1, "Alpha Link", "https://example.org/partial?q=1", "anchor");
                }
            });
        }

        [WordComFact]
        public void PlainTextInsideOneTableCellCanBeReplacedAndUndone()
        {
            RunInSta(word =>
            {
                object table = word.CreateTwoCellTable("Cell text", "Keep neighboring cell");
                object range = word.Find("Cell text"); // Ordinary text selection, without cell markers.
                using (var editor = WordSelectionEditor.Capture(word.Document, (object)range))
                {
                    Assert.True(editor.IsCurrent);
                    Assert.Equal("Updated cell text", editor.Apply("Updated cell text", "Cell text"));
                    word.AssertCells(table, "Updated cell text", "Keep neighboring cell");
                    word.AssertSelectionMatches((object)range, "Updated cell text");
                    Assert.True((bool)((dynamic)word.Document).Undo(1));
                    word.AssertCells(table, "Cell text", "Keep neighboring cell");
                }
            });
        }

        [WordComFact]
        public void WholeSingleCellSelectionIncludingEndOfCellCanBeEditedWithoutChangingTable()
        {
            RunInSta(word =>
            {
                object table = word.CreateTwoCellTable("Cell text", "Keep neighboring cell");
                object range = word.CellRange(table, 1);
                Assert.Equal("Cell text\r\a", (string)((dynamic)range).Text);
                // Exercise the same preprocessing as the real Outlook controller, not a test-only trim.
                Assert.Equal("Cell text", OutlookTextActionController.AdjustAndReadSelection(range));
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    editor.Apply(editor.Plan.ModelText.Replace("Cell text", "Updated cell text"), "Whole cell");
                    word.AssertCells(table, "Updated cell text", "Keep neighboring cell");
                    dynamic selection = word.Selection();
                    dynamic cellRange = word.CellRange(table, 1);
                    Assert.Equal((int)cellRange.Start, (int)selection.Start);
                    Assert.InRange((int)selection.End, (int)cellRange.End - 2, (int)cellRange.End);
                    Assert.StartsWith("Updated cell text", (string)selection.Text);
                    Assert.True((bool)((dynamic)word.Document).Undo(1));
                    word.AssertCells(table, "Cell text", "Keep neighboring cell");
                }
            });
        }

        [WordComFact]
        public void SelectionAcrossTwoCellsIsRejectedWithoutChangingDocument()
        {
            RunInSta(word =>
            {
                object table = word.CreateTwoCellTable("First cell", "Second cell");
                object range = word.Own(((dynamic)table).Range);
                string before = word.Text;
                Assert.Throws<InvalidOperationException>(() =>
                {
                    using (WordSelectionEditor.Capture(word.Document, range)) { }
                });
                Assert.Equal(before, word.Text);
                word.AssertCells(table, "First cell", "Second cell");
            });
        }

        [WordComFact]
        public void CallerDetectsTextChangedWhileModelWasRunningBeforeApplyingReplacement()
        {
            RunInSta(word =>
            {
                word.SetText("Original selected text and untouched suffix");
                object range = word.Find("Original selected text");
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    ((dynamic)word.Find("selected")).Text = "manually edited";
                    // The Outlook controller checks IsCurrent before Apply; Apply does not own that guard.
                    Assert.False(editor.IsCurrent);
                    Assert.Equal("Original manually edited text and untouched suffix\r", word.Text);
                }
            });
        }

        [WordComFact]
        public void DelayedUndoRefusesToUndoAnotherEditOutsideTheSelection()
        {
            RunInSta(word =>
            {
                word.SetText("Original selection and outside text");
                object range = word.Find("Original selection");
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    editor.Apply("AI replacement", "First edit");
                    ((dynamic)word.Find("outside text")).Text = "later manual edit";
                    Assert.False(editor.Undo());
                    Assert.Equal("AI replacement and later manual edit\r", word.Text);
                }
            });
        }

        [WordComFact]
        public void FormattingOnlyChangeInvalidatesTheCapturedSelection()
        {
            RunInSta(word =>
            {
                word.SetText("Original selected text");
                object range = word.BodyWithoutFinalParagraph();
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    word.SetEmphasis("selected", AccentColor);
                    Assert.False(editor.IsCurrent);
                }
            });
        }

        [WordComFact]
        public void EditorUndoRestoresTheOriginalRangeWhenNoLaterEditsExist()
        {
            RunInSta(word =>
            {
                word.SetText("Original selected text");
                object range = word.BodyWithoutFinalParagraph();
                using (var editor = WordSelectionEditor.Capture(word.Document, range))
                {
                    Assert.True(editor.IsCurrent);
                    editor.Apply("Updated selected text", "Undo button");
                    Assert.True(editor.Undo());
                    Assert.Equal("Original selected text\r", word.Text);
                    word.AssertSelectionMatches(range, "Original selected text");
                }
            });
        }

        private static void RunInSta(Action<WordDocument> test)
        {
            ExceptionDispatchInfo failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    using (var word = new WordDocument()) test(word);
                }
                catch (Exception error) { failure = ExceptionDispatchInfo.Capture(error); }
            }) { IsBackground = true, Name = "OutlookAI isolated Word COM test" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(TimeSpan.FromSeconds(90)))
                throw new TimeoutException("Word COM test exceeded 90 seconds. Check Word first-run or modal dialogs. "
                    + "No Word process was killed; the private instance is closed by finally when its COM call returns.");
            failure?.Throw();
        }

        private sealed class WordDocument : IDisposable
        {
            private readonly List<object> _owned = new List<object>();
            private object _application;
            private object _documents;
            private object _document;

            public WordDocument()
            {
                try
                {
                    Type type = Type.GetTypeFromProgID("Word.Application", true);
                    _application = Activator.CreateInstance(type);
                    dynamic application = _application;
                    application.Visible = false;
                    application.DisplayAlerts = 0; // wdAlertsNone; all document close calls also explicitly discard changes.
                    _documents = application.Documents;
                    _document = ((dynamic)_documents).Add(); // New, unsaved document only. Never GetActiveObject/Open.
                    ((dynamic)_document).TrackRevisions = false;
                }
                catch { Dispose(); throw; }
            }

            public object Document => _document;
            public string Text => (string)((dynamic)Own(((dynamic)_document).Content)).Text;
            public object Own(object value) { _owned.Add(value); return value; }
            public object Range(int start, int end) => Own(((dynamic)_document).Range(start, end));
            public object Selection() => Own(((dynamic)_application).Selection);

            public void SetText(string text) { ((dynamic)Own(((dynamic)_document).Content)).Text = text; }

            public object BodyWithoutFinalParagraph()
            {
                dynamic content = Own(((dynamic)_document).Content);
                content.SetRange((int)content.Start, (int)content.End - 1);
                return content;
            }

            public object Find(string text)
            {
                dynamic range = Own(((dynamic)_document).Content);
                dynamic find = Own(range.Find);
                Assert.True((bool)find.Execute(FindText: text, MatchCase: true, MatchWholeWord: false,
                    MatchWildcards: false, Forward: true, Wrap: 0, Format: false), "Missing fixture text: " + text);
                return range;
            }

            public void AddLink(string caption, string address, string subAddress)
            {
                dynamic links = Own(((dynamic)_document).Hyperlinks);
                Own(links.Add(Anchor: Find(caption), Address: address, SubAddress: subAddress));
            }

            public void SetEmphasis(string text, int color)
            {
                dynamic font = Own(((dynamic)Find(text)).Font);
                font.Bold = -1;
                font.Color = color;
            }

            public void AssertEmphasis(string text, int color)
            {
                dynamic font = Own(((dynamic)Find(text)).Font);
                Assert.Equal(-1, (int)font.Bold);
                Assert.Equal(color, (int)font.Color);
            }

            public void AssertLinks()
            {
                dynamic fields = Own(((dynamic)_document).Fields);
                dynamic links = Own(((dynamic)_document).Hyperlinks);
                Assert.Equal(2, (int)fields.Count);
                Assert.Equal(2, (int)links.Count);
                AssertLink(1, "First link", "https://example.org/first?q=1", "section-one");
                AssertLink(2, "Second link", "https://example.org/second?q=2", "section-two");
            }

            public void AssertLink(int index, string caption, string address, string subAddress)
            {
                dynamic links = Own(((dynamic)_document).Hyperlinks);
                dynamic link = Own(links[index]);
                Assert.Equal(caption, (string)link.TextToDisplay);
                Assert.Equal(address, (string)link.Address);
                Assert.Equal(subAddress, (string)link.SubAddress);
            }

            public void AssertSelectionMatches(object range, string text)
            {
                dynamic actual = Selection();
                Assert.Equal((int)((dynamic)range).Start, (int)actual.Start);
                Assert.Equal((int)((dynamic)range).End, (int)actual.End);
                Assert.Equal(text, (string)actual.Text);
            }

            public object CreateTwoCellTable(string first, string second)
            {
                dynamic tables = Own(((dynamic)_document).Tables);
                object table = Own(tables.Add(Range(0, 0), 1, 2));
                ((dynamic)CellRange(table, 1)).Text = first;
                ((dynamic)CellRange(table, 2)).Text = second;
                return table;
            }

            public object CellRange(object table, int column)
            {
                dynamic cell = Own(((dynamic)table).Cell(1, column));
                return Own(cell.Range);
            }

            public void AssertCells(object table, string first, string second)
            {
                dynamic rows = Own(((dynamic)table).Rows);
                dynamic columns = Own(((dynamic)table).Columns);
                Assert.Equal(1, (int)rows.Count);
                Assert.Equal(2, (int)columns.Count);
                Assert.Equal(first + "\r\a", (string)((dynamic)CellRange(table, 1)).Text);
                Assert.Equal(second + "\r\a", (string)((dynamic)CellRange(table, 2)).Text);
            }

            public void Dispose()
            {
                // Release subordinate RCWs before closing the private document and application.
                for (int i = _owned.Count - 1; i >= 0; i--) Release(_owned[i]);
                _owned.Clear();
                try
                {
                    if (_document != null) ((dynamic)_document).Close(0); // wdDoNotSaveChanges
                }
                finally
                {
                    Release(_document);
                    _document = null;
                    try
                    {
                        // Do not close documents that somebody might have opened in this instance during the test.
                        if (_application != null && (_documents == null || (int)((dynamic)_documents).Count == 0))
                            ((dynamic)_application).Quit(0); // wdDoNotSaveChanges
                    }
                    finally
                    {
                        Release(_documents);
                        Release(_application);
                        _documents = null;
                        _application = null;
                    }
                }
            }

            private static void Release(object value)
            {
                if (value == null || !Marshal.IsComObject(value)) return;
                try { Marshal.ReleaseComObject(value); }
                catch (InvalidComObjectException)
                {
                    // An already-disconnected subordinate RCW must not prevent closing our document.
                }
            }
        }
    }
}
