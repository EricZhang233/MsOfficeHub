using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.FileProperties;

namespace MsOfficeHub
{
    public sealed partial class MainWindow : Window
    {
        public ObservableCollection<OfficeAppItem> OfficeApps { get; } = new();

        public MainWindow()
        {
            InitializeComponent();
            InitializeTitleBar();
            _ = LoadOfficeAppsAsync();
        }

        private void InitializeTitleBar()
        {
            SystemBackdrop = new DesktopAcrylicBackdrop();
            ExtendsContentIntoTitleBar = true;
            SetTitleBar(TitleBarGrid);
            Title = "MsOfficeHub";
            VersionText.Text = AppVersion.DisplayVersion;

            var titleBar = AppWindow.TitleBar;
            titleBar.ButtonBackgroundColor = Windows.UI.Color.FromArgb(0, 0, 0, 0);
            titleBar.ButtonInactiveBackgroundColor = Windows.UI.Color.FromArgb(0, 0, 0, 0);

            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using var stream = assembly.GetManifestResourceStream("MsOfficeHub.ico.ico");
                if (stream != null)
                {
                    var tempIconPath = Path.Combine(Path.GetTempPath(), "MsOfficeHub_TempIcon.ico");
                    using var fileStream = new FileStream(tempIconPath, FileMode.Create, FileAccess.Write);
                    stream.CopyTo(fileStream);
                    fileStream.Close();

                    AppWindow.SetIcon(tempIconPath);
                    TitleBarIcon.Source = new BitmapImage(new Uri(tempIconPath));
                }
            }
            catch
            {
            }
        }

        private async Task LoadOfficeAppsAsync()
        {
            OfficeApps.Clear();
            var definitions = new[]
            {
                new OfficeAppDefinition("Word", new[] { "WINWORD.EXE" }, Array.Empty<string>(), "word.png"),
                new OfficeAppDefinition("Excel", new[] { "EXCEL.EXE" }, Array.Empty<string>(), "excel.png"),
                new OfficeAppDefinition("PowerPoint", new[] { "POWERPNT.EXE" }, Array.Empty<string>(), "powerpoint.png"),
                new OfficeAppDefinition("Outlook", new[] { "OUTLOOK.EXE" }, Array.Empty<string>(), "outlook.png"),
                new OfficeAppDefinition("OneNote", new[] { "ONENOTE.EXE" }, Array.Empty<string>(), "onenote.png"),
                new OfficeAppDefinition("Access", new[] { "MSACCESS.EXE" }, Array.Empty<string>(), "access.png"),
                new OfficeAppDefinition("Publisher", new[] { "MSPUB.EXE" }, Array.Empty<string>(), "publisher.png"),
                new OfficeAppDefinition("Teams", new[] { "TEAMS.EXE", "MS-TEAMS.EXE" }, new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Teams", "current", "Teams.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Teams", "current", "ms-teams.exe")
                }, "teams.png"),
                new OfficeAppDefinition("OneDrive", new[] { "ONEDRIVE.EXE" }, new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "OneDrive", "OneDrive.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft OneDrive", "OneDrive.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft OneDrive", "OneDrive.exe")
                }, "onedrive.png"),
                new OfficeAppDefinition("Microsoft 365 Copilot", new[] { "M365COPILOT.EXE", "COPILOTOFFICE.EXE" }, Array.Empty<string>(), "copilot.png"),
                new OfficeAppDefinition("Clipchamp", new[] { "CLIPCHAMP.EXE" }, Array.Empty<string>(), "clipchamp.png"),
                new OfficeAppDefinition("Project", new[] { "WINPROJ.EXE" }, Array.Empty<string>(), "project.png"),
                new OfficeAppDefinition("Visio", new[] { "VISIO.EXE" }, Array.Empty<string>(), "visio.png"),
                new OfficeAppDefinition("Skype for Business", new[] { "LYNC.EXE" }, Array.Empty<string>(), "skype.png")
            };

            var detector = new OfficeDetector();

            foreach (var definition in definitions)
            {
                var detection = detector.Detect(definition.ExeNames, definition.ExtraPaths, ConfigService.ForceShowAllApps);
                if (!detection.IsInstalled && !ConfigService.ForceShowAllApps)
                {
                    continue;
                }

                var fallbackIcon = TryLoadFallbackIcon(definition.FallbackFileName);
                var item = new OfficeAppItem(definition.Name, fallbackIcon);
                OfficeApps.Add(item);

                if (!string.IsNullOrWhiteSpace(detection.InstallPath))
                {
                    var icon = await TryLoadFileIconAsync(detection.InstallPath);
                    if (icon != null)
                    {
                        item.Icon = icon;
                    }
                }
            }
        }

        private ImageSource? TryLoadFallbackIcon(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            var path = Path.Combine(AppContext.BaseDirectory, "Assets", "Office", fileName);
            if (!File.Exists(path))
            {
                return null;
            }

            return new BitmapImage(new Uri(path));
        }

        private async Task<ImageSource?> TryLoadFileIconAsync(string path)
        {
            try
            {
                var file = await StorageFile.GetFileFromPathAsync(path);
                using var thumbnail = await file.GetThumbnailAsync(ThumbnailMode.SingleItem, 64);
                if (thumbnail == null || thumbnail.Size == 0)
                {
                    return null;
                }

                var image = new BitmapImage();
                await image.SetSourceAsync(thumbnail);
                return image;
            }
            catch
            {
                return null;
            }
        }

    }

    public sealed class OfficeAppItem : INotifyPropertyChanged
    {
        private ImageSource? _icon;

        public OfficeAppItem(string displayName, ImageSource? icon)
        {
            DisplayName = displayName;
            _icon = icon;
        }

        public string DisplayName { get; }
        public ImageSource? Icon
        {
            get => _icon;
            set
            {
                if (!Equals(_icon, value))
                {
                    _icon = value;
                    OnPropertyChanged(nameof(Icon));
                    OnPropertyChanged(nameof(ImageVisibility));
                }
            }
        }

        public Visibility ImageVisibility => Icon == null ? Visibility.Collapsed : Visibility.Visible;

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal sealed class OfficeAppDefinition
    {
        public OfficeAppDefinition(string name, string[] exeNames, string[] extraPaths, string fallbackFileName)
        {
            Name = name;
            ExeNames = exeNames;
            ExtraPaths = extraPaths;
            FallbackFileName = fallbackFileName;
        }

        public string Name { get; }
        public string[] ExeNames { get; }
        public string[] ExtraPaths { get; }
        public string FallbackFileName { get; }
    }
}
