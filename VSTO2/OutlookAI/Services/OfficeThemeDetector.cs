using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    internal static class OfficeThemeDetector
    {
        public static string GetThemeName()
        {
            if (SystemInformation.HighContrast)
            {
                return "high-contrast";
            }

            var officeTheme = ReadDword(
                @"Software\Microsoft\Office\16.0\Common",
                "UI Theme");

            // Office: 1 = dark gray, 2 = black, 3 = white,
            // 0 = colorful, 4 = follow system.
            if (officeTheme == 1 || officeTheme == 2)
            {
                return "dark";
            }
            if (officeTheme == 0 || officeTheme == 3)
            {
                return "light";
            }

            var appsUseLightTheme = ReadDword(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                "AppsUseLightTheme");
            return appsUseLightTheme == 0 ? "dark" : "light";
        }

        public static Color GetBackgroundColor()
        {
            var theme = GetThemeName();
            if (theme == "high-contrast") return Color.Black;
            if (theme == "dark") return Color.FromArgb(31, 31, 31);
            return Color.FromArgb(247, 249, 252);
        }

        private static int? ReadDword(string path, string name)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(path))
                {
                    var value = key?.GetValue(name);
                    return value == null ? (int?)null : Convert.ToInt32(value);
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
