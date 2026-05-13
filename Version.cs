namespace MsOfficeHub
{
    public static class AppVersion
    {
        public const string Version = "1.0.0.0";
        public static string DisplayVersion
        {
            get
            {
                var parts = Version.Split('.');
                if (parts.Length >= 4 && int.TryParse(parts[3], out var build) && build != 0)
                {
                    return $"Ver.{Version}";
                }
                var displayVer = string.Join(".", parts.AsSpan(0, 3));
                return $"Ver.{displayVer}";
            }
        }
    }
}
