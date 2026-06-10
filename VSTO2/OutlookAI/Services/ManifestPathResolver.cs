using System;
using System.IO;

namespace OutlookAI.Services
{
    internal static class ManifestPathResolver
    {
        public static string ResolveProjectManifest(string relativePath)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var deployed = Path.Combine(baseDir, relativePath);
            if (File.Exists(deployed))
            {
                return deployed;
            }

            var dir = new DirectoryInfo(baseDir);
            while (dir != null)
            {
                var candidate = Path.Combine(dir.FullName, relativePath);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
                dir = dir.Parent;
            }

            return deployed;
        }

        public static string AppDataFile(string filename)
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OutlookAI");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, filename);
        }

        public static string LocalAppDataFile(string filename)
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, filename);
        }
    }
}
