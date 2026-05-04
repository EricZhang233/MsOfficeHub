using Microsoft.Win32;
using System;
using System.IO;

namespace MsOfficeHub
{
    internal sealed class OfficeDetectionResult
    {
        public OfficeDetectionResult(bool isInstalled, string? installPath)
        {
            IsInstalled = isInstalled;
            InstallPath = installPath;
        }

        public bool IsInstalled { get; }
        public string? InstallPath { get; }
    }

    internal sealed class OfficeDetector
    {
        public OfficeDetectionResult Detect(string[] exeNames, string[] extraPaths, bool forceShowAllApps)
        {
            if (forceShowAllApps)
            {
                return new OfficeDetectionResult(true, null);
            }

            foreach (var path in extraPaths)
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    return new OfficeDetectionResult(true, path);
                }
            }

            foreach (var exeName in exeNames)
            {
                if (TryResolveAppPath(exeName, out var resolvedPath) && File.Exists(resolvedPath))
                {
                    return new OfficeDetectionResult(true, resolvedPath);
                }
            }

            var officePath = TryResolveOfficePath(exeNames);
            if (!string.IsNullOrWhiteSpace(officePath))
            {
                return new OfficeDetectionResult(true, officePath);
            }

            return new OfficeDetectionResult(false, null);
        }

        private bool TryResolveAppPath(string exeName, out string path)
        {
            path = string.Empty;

            if (string.IsNullOrWhiteSpace(exeName))
            {
                return false;
            }

            if (TryResolveAppPath(RegistryView.Registry64, exeName, out path))
            {
                return true;
            }

            if (TryResolveAppPath(RegistryView.Registry32, exeName, out path))
            {
                return true;
            }

            return false;
        }

        private bool TryResolveAppPath(RegistryView view, string exeName, out string path)
        {
            path = string.Empty;

            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            using var appPathsKey = baseKey.OpenSubKey($"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\{exeName}");

            if (appPathsKey?.GetValue("") is string appPath && !string.IsNullOrWhiteSpace(appPath))
            {
                path = appPath.Trim();
                return true;
            }

            return false;
        }

        private string? TryResolveOfficePath(string[] exeNames)
        {
            var officeRoot = GetOfficeInstallRoot();
            if (!string.IsNullOrWhiteSpace(officeRoot))
            {
                foreach (var exeName in exeNames)
                {
                    if (string.IsNullOrWhiteSpace(exeName))
                    {
                        continue;
                    }

                    var candidate = Path.Combine(officeRoot, exeName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            var fallbackRoots = GetOfficeFallbackRoots();
            foreach (var root in fallbackRoots)
            {
                if (string.IsNullOrWhiteSpace(root))
                {
                    continue;
                }

                foreach (var exeName in exeNames)
                {
                    if (string.IsNullOrWhiteSpace(exeName))
                    {
                        continue;
                    }

                    var candidate = Path.Combine(root, exeName);
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private string? GetOfficeInstallRoot()
        {
            var clickToRunPath = TryGetRegistryValue(RegistryHive.LocalMachine, RegistryView.Registry64, "SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration", "InstallationPath");
            if (!string.IsNullOrWhiteSpace(clickToRunPath))
            {
                var officeRoot = Path.Combine(clickToRunPath, "Office16");
                if (Directory.Exists(officeRoot))
                {
                    return officeRoot;
                }
            }

            var rootPath = TryGetRegistryValue(RegistryHive.LocalMachine, RegistryView.Registry64, "SOFTWARE\\Microsoft\\Office\\16.0\\Common\\InstallRoot", "Path");
            if (!string.IsNullOrWhiteSpace(rootPath) && Directory.Exists(rootPath))
            {
                return rootPath;
            }

            rootPath = TryGetRegistryValue(RegistryHive.LocalMachine, RegistryView.Registry32, "SOFTWARE\\Microsoft\\Office\\16.0\\Common\\InstallRoot", "Path");
            if (!string.IsNullOrWhiteSpace(rootPath) && Directory.Exists(rootPath))
            {
                return rootPath;
            }

            return null;
        }

        private string[] GetOfficeFallbackRoots()
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            return new[]
            {
                Path.Combine(programFiles, "Microsoft Office", "root", "Office16"),
                Path.Combine(programFilesX86, "Microsoft Office", "root", "Office16"),
                Path.Combine(programFiles, "Microsoft Office", "Office16"),
                Path.Combine(programFilesX86, "Microsoft Office", "Office16")
            };
        }

        private string? TryGetRegistryValue(RegistryHive hive, RegistryView view, string subKey, string valueName)
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                using var key = baseKey.OpenSubKey(subKey);
                if (key?.GetValue(valueName) is string value && !string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }
            catch
            {
                return null;
            }

            return null;
        }
    }
}
