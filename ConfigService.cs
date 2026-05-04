using System;
using System.IO;
using System.Text.Json;

namespace MsOfficeHub
{
    internal static class ConfigService
    {
        private static bool _initialized;

        public static bool ForceShowAllApps { get; private set; }
        public static string DebugIniPath { get; private set; } = string.Empty;
        public static string ConfigJsonPath { get; private set; } = string.Empty;

        public static void Initialize()
        {
            if (_initialized)
            {
                return;
            }

            _initialized = true;

            var baseDir = AppContext.BaseDirectory;
            DebugIniPath = Path.Combine(baseDir, "debug.ini");
            ConfigJsonPath = Path.Combine(baseDir, "config.json");

            ForceShowAllApps = ReadForceShowAllApps(DebugIniPath);
            _ = TryLoadConfigJson(ConfigJsonPath);
        }

        private static bool ReadForceShowAllApps(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    return false;
                }

                foreach (var rawLine in File.ReadAllLines(path))
                {
                    var line = rawLine.Trim();
                    if (line.Length == 0 || line.StartsWith("#") || line.StartsWith(";"))
                    {
                        continue;
                    }

                    var parts = line.Split('=', 2, StringSplitOptions.TrimEntries);
                    if (parts.Length != 2)
                    {
                        continue;
                    }

                    if (!parts[0].Equals("ForceShowAllApps", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var valueText = parts[1].Trim();
                    if (valueText.Equals("true", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    if (valueText.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    if (int.TryParse(valueText, out var value))
                    {
                        if (value == 1)
                        {
                            return true;
                        }

                        if (value == 0)
                        {
                            return false;
                        }

                        return false;
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static bool TryLoadConfigJson(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    return false;
                }

                using var doc = JsonDocument.Parse(File.ReadAllText(path));
                return doc.RootElement.ValueKind == JsonValueKind.Object;
            }
            catch
            {
                return false;
            }
        }
    }
}
