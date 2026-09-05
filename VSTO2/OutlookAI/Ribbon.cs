using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OutlookAI.Diagnostics;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            // Outlook calls GetCustomUI once per ribbon context. We serve
            // the same customUI XML to mail, appointment and task compose
            // Inspectors plus Explorer.
            // Tab definitions that don't match the current context are
            // silently ignored by Office's ribbon-XML applier, so a single
            // XML payload safely covers both cases.
            if (ribbonID == "Microsoft.Outlook.Mail.Compose" ||
                ribbonID == "Microsoft.Outlook.Mail.Read" ||
                ribbonID == "Microsoft.Outlook.MeetingRequest.Read" ||
                ribbonID == "Microsoft.Outlook.Explorer" ||
                ribbonID == "Microsoft.Outlook.Appointment" ||
                ribbonID == "Microsoft.Outlook.Task")
            {
                return GetResourceText("OutlookAI.Ribbon.xml");
            }
            return null;
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnAIAssistantClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowTaskPane();
        }

        public void OnReportsClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowReportsTaskPane();
        }

        public void OnSettingsClick(Office.IRibbonControl control)
        {
            using (var settingsForm = new SettingsForm())
            {
                settingsForm.ShowDialog();
            }
        }

        public bool GetSelectionMenuVisible(Office.IRibbonControl control)
        {
            try
            {
                var controller = Globals.ThisAddIn?.TextActionController;
                return controller != null && controller.CanExecute(GetInspector(control));
            }
            catch (Exception ex)
            {
                TraceLog.Write("Selection menu visibility error: " + ex, "Ribbon");
                return false;
            }
        }

        public string GetSelectionMenuContent(Office.IRibbonControl control)
        {
            try
            {
                var controller = Globals.ThisAddIn?.TextActionController;
                return controller != null
                    ? controller.GetMenuContent(GetInspector(control))
                    : null;
            }
            catch (Exception ex)
            {
                TraceLog.Write("Selection menu content error: " + ex, "Ribbon");
                return null;
            }
        }

        public async void OnTextActionClick(Office.IRibbonControl control)
        {
            try
            {
                var controller = Globals.ThisAddIn?.TextActionController;
                if (controller == null || string.IsNullOrWhiteSpace(control?.Tag)) return;
                await controller.ExecuteAsync(GetInspector(control), control.Tag);
            }
            catch (Exception ex)
            {
                TraceLog.Write("Selection action error: " + ex, "Ribbon");
                MessageBox.Show(
                    "Не удалось обработать выделенный текст: " + ex.Message,
                    "AI-ассистент",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static Outlook.Inspector GetInspector(Office.IRibbonControl control)
        {
            var contextInspector = control?.Context as Outlook.Inspector;
            if (contextInspector != null) return contextInspector;
            try { return Globals.ThisAddIn?.Application?.ActiveInspector(); }
            catch { return null; }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            foreach (string name in resourceNames)
            {
                if (name.EndsWith(resourceName, StringComparison.OrdinalIgnoreCase))
                {
                    var stream = asm.GetManifestResourceStream(name);
                    if (stream == null) continue;
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }
    }
}
