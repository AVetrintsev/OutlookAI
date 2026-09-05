using System.IO;
using Xunit;

namespace OutlookAI.Tests.TaskPane.Chat
{
    public sealed class InlineTextActionsWebUiTests
    {
        [Fact]
        public void EditorSelectionDialog_UsesFixedSourceOutputAndSurface()
        {
            var html = ReadWebUiFile("index.html");
            var script = ReadWebUiFile("chat.js");

            Assert.Contains("value=\"selected_text\"", html);
            Assert.Contains("value=\"replace_selection\"", html);
            Assert.Contains("value=\"task\"", html);
            Assert.Contains("editingActionGroupId === 'text_editing'", script);
            Assert.Contains("$customActionSource.value = 'selected_text';", script);
            Assert.Contains("$customActionOutput.value = 'replace_selection';", script);
            Assert.Contains("surface: isEditorSelection ? 'editor_selection' : 'assistant'", script);
            Assert.Contains("output: isEditorSelection", script);
            Assert.Contains("? 'replace_selection'", script);
        }

        [Fact]
        public void EditorSelectionDialog_HidesIrrelevantContextAndToolFields()
        {
            var script = ReadWebUiFile("chat.js");

            Assert.Contains("setFieldVisible($customActionSourceField, !isEditorSelection);", script);
            Assert.Contains("setFieldVisible($customActionFullBodiesField, !isEditorSelection);", script);
            Assert.Contains("setFieldVisible($customActionOutputField, !isEditorSelection);", script);
            Assert.Contains("!isEditorSelection && customActionToolCatalog.length > 0", script);
            Assert.Contains("var allowedTools = isEditorSelection ? []", script);
        }

        [Fact]
        public void TextEditingGroup_KeepsExistingCrudReorderAndResetControls()
        {
            var script = ReadWebUiFile("chat.js");

            Assert.Contains("openCustomActionDialog({ group_id: editingActionGroupId });", script);
            Assert.Contains("openCustomActionDialog(action);", script);
            Assert.Contains("type: 'custom_action_delete'", script);
            Assert.Contains("type: 'custom_action_reorder'", script);
            Assert.Contains("type: 'custom_action_reset_group'", script);
        }

        private static string ReadWebUiFile(string fileName)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(current.FullName, "OutlookAI", "WebUI", fileName);
                if (File.Exists(candidate)) return File.ReadAllText(candidate);
                current = current.Parent;
            }

            throw new FileNotFoundException("Could not find WebUI/" + fileName);
        }
    }
}
