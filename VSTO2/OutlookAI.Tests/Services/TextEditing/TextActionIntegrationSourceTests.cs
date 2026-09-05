using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OutlookAI.Tests.Services.TextEditing
{
    public sealed class TextActionIntegrationSourceTests
    {
        [Fact]
        public void RibbonXml_RegistersSelectionDynamicMenuAndEditorTabs()
        {
            var document = XDocument.Load(FindSourceFile("OutlookAI", "Ribbon.xml"));

            var contextMenu = document
                .Descendants()
                .Single(element => element.Name.LocalName == "contextMenu"
                    && (string)element.Attribute("idMso") == "ContextMenuText");
            var dynamicMenu = contextMenu
                .Elements()
                .Single(element => element.Name.LocalName == "dynamicMenu");

            Assert.Equal("GetSelectionMenuVisible", (string)dynamicMenu.Attribute("getVisible"));
            Assert.Equal("GetSelectionMenuContent", (string)dynamicMenu.Attribute("getContent"));
            Assert.Equal("true", (string)dynamicMenu.Attribute("invalidateContentOnDrop"));

            var tabIds = document
                .Descendants()
                .Where(element => element.Name.LocalName == "tab")
                .Select(element => (string)element.Attribute("idMso"))
                .ToArray();
            Assert.Contains("TabAppointment", tabIds);
            Assert.Contains("TabTask", tabIds);
        }

        [Fact]
        public void RibbonSource_ServesComposeSurfacesAndForwardsActionTag()
        {
            var source = ReadSource("OutlookAI", "Ribbon.cs");

            Assert.Contains("Microsoft.Outlook.Mail.Compose", source);
            Assert.Contains("Microsoft.Outlook.Appointment", source);
            Assert.Contains("Microsoft.Outlook.Task", source);
            Assert.Contains("public async void OnTextActionClick", source);
            Assert.Contains(
                "controller.ExecuteAsync(GetInspector(control), control.Tag)",
                source);
        }

        [Fact]
        public void ThisAddInSource_OwnsControllerLifetime()
        {
            var source = ReadSource("OutlookAI", "ThisAddIn.cs");

            Assert.Contains("TextActionController = new OutlookTextActionController(", source);
            Assert.Contains("TextActionController?.Dispose()", source);
        }

        [Fact]
        public void ChatControllerSource_RoutesTextActionsBeforeCustomActionRunner()
        {
            var source = ReadSource(
                "OutlookAI",
                "TaskPane",
                "Chat",
                "ChatController.cs");
            var textActionRoute = source.IndexOf(
                "TextActionCatalog.IsTextAction(action)",
                StringComparison.Ordinal);
            var genericRunner = source.IndexOf(
                "new CustomActionRunner(_chat, _surface, _toolHost,",
                StringComparison.Ordinal);

            Assert.True(textActionRoute >= 0, "Text action routing branch was not found.");
            Assert.True(genericRunner >= 0, "Generic custom-action runner was not found.");
            Assert.True(
                textActionRoute < genericRunner,
                "Text actions must be routed before the generic CustomActionRunner path.");
            Assert.Contains("await textController.ExecuteAsync(null, action.Id, _skillUi.PinnedIds);", source);
        }

        [Fact]
        public void LiveOutlookSurfaceSource_HandlesMeetingAndTaskComposeItems()
        {
            var source = ReadSource(
                "OutlookAI",
                "Services",
                "Tools",
                "LiveOutlookSurface.cs");

            Assert.Contains("currentItem is Outlook.AppointmentItem appointment", source);
            Assert.Contains("currentItem is Outlook.TaskItem task", source);
            Assert.Contains("ItemType = \"meeting\"", source);
            Assert.Contains("ItemType = \"task\"", source);
        }

        private static string ReadSource(params string[] relativeParts)
        {
            return File.ReadAllText(FindSourceFile(relativeParts));
        }

        private static string FindSourceFile(params string[] relativeParts)
        {
            var current = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (current != null)
            {
                var candidate = Path.Combine(
                    current.FullName,
                    Path.Combine(relativeParts));
                if (File.Exists(candidate)) return candidate;

                current = current.Parent;
            }

            throw new FileNotFoundException(
                "Could not find source file.",
                Path.Combine(relativeParts));
        }
    }
}
