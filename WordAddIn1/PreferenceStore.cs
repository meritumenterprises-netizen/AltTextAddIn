using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace WordAddIn1
{
    public static class PreferenceStore
    {
        private static readonly string FolderPath =
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "PiotrLuczak",
                "AltTextAddIn");

        private static readonly string FilePath =
            Path.Combine(FolderPath, "preferences.txt");

        public static void SavePreference(string value)
        {
            Directory.CreateDirectory(FolderPath);
            File.WriteAllText(FilePath, value ?? string.Empty);
        }

        public static bool GetDisplayOnStartup()
        {
            var preference = LoadPreference();
            return preference == "Visible" || preference == string.Empty;
        }

        public static void SetDisplayOnStartup(bool displayOnStartup)
        {
            SavePreference(displayOnStartup ? "Visible" : "Hidden");
        }
        private static string LoadPreference()
        {
            return File.Exists(FilePath)
                ? File.ReadAllText(FilePath)
                : string.Empty;
        }
    }
}
