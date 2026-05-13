using Microsoft.UI.Xaml;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Input;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.FileProperties;

namespace MsOfficeHub
{
    public sealed partial class MainWindow : Window
    {
        public ObservableCollection<OfficeAppItem> OfficeApps { get; } = new();
        public ObservableCollection<RecentItem> RecentItems { get; } = new();
        private readonly DispatcherQueueTimer _recentTimer;
        private DateTimeOffset _lastRecentRefresh = DateTimeOffset.MinValue;
        private static readonly TimeSpan RecentRefreshMinInterval = TimeSpan.FromSeconds(10);
        private const string RegistryPath = @"Software\EricSoft\MsOfficeHub";

        public MainWindow()
        {
            InitializeComponent();
            LoadWindowState();
            InitializeTitleBar();
            _ = LoadOfficeAppsAsync();
            _ = RefreshRecentAsync();

            Activated += OnWindowActivated;

            _recentTimer = DispatcherQueue.GetForCurrentThread().CreateTimer();
            _recentTimer.Interval = TimeSpan.FromSeconds(60);
            _recentTimer.Tick += async (_, _) => await RefreshRecentAsync();
            _recentTimer.Start();
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
                new OfficeAppDefinition("Outlook", new[] { "OUTLOOK.EXE" }, new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WindowsApps", "olk.exe")
                }, "outlook.png"),
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

                var launchPath = detection.InstallPath;
                if (string.IsNullOrWhiteSpace(launchPath) && definition.ExeNames.Length > 0)
                {
                    launchPath = definition.ExeNames[0];
                }

                var fallbackIcon = await TryLoadFallbackIconAsync(definition.FallbackFileName);
                var item = new OfficeAppItem(definition.Name, fallbackIcon, launchPath);
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

        private async Task<ImageSource?> TryLoadFallbackIconAsync(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            var resourceName = $"Assets.Office.{fileName}";
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using var stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null)
            {
                return null;
            }

            var image = new BitmapImage();
            await image.SetSourceAsync(stream.AsRandomAccessStream());
            return image;
        }

        private async Task<ImageSource?> TryLoadFileIconAsync(string path)
        {
            try
            {
                var file = await Windows.Storage.StorageFile.GetFileFromPathAsync(path);
                using var thumbnail = await file.GetThumbnailAsync(Windows.Storage.FileProperties.ThumbnailMode.SingleItem, 64);
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

        [System.Runtime.InteropServices.DllImport("shell32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        private const uint SHGFI_ICON = 0x000000100;
        private const uint SHGFI_LARGEICON = 0x000000000;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        private readonly System.Collections.Generic.Dictionary<string, ImageSource> _fileIconCache = new();
        private readonly System.Collections.Generic.Dictionary<string, ImageSource> _extensionIconCache = new(StringComparer.OrdinalIgnoreCase);

        private async Task<ImageSource?> GetIconForExtensionAsync(string extension)
        {
            if (string.IsNullOrEmpty(extension)) return null;
            if (!extension.StartsWith(".")) extension = "." + extension;

            if (_extensionIconCache.TryGetValue(extension, out var cached) && cached != null)
            {
                return cached;
            }

            try
            {
                var shinfo = new SHFILEINFO();
                var cbFileInfo = (uint)System.Runtime.InteropServices.Marshal.SizeOf(shinfo);
                SHGetFileInfo(extension, FILE_ATTRIBUTE_NORMAL, ref shinfo, cbFileInfo, SHGFI_ICON | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES);

                if (shinfo.hIcon != IntPtr.Zero)
                {
                    using var icon = System.Drawing.Icon.FromHandle(shinfo.hIcon);
                    using var bmp = icon.ToBitmap();
                    using var ms = new System.IO.MemoryStream();
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;

                    var image = new BitmapImage();
                    await image.SetSourceAsync(ms.AsRandomAccessStream());

                    DestroyIcon(shinfo.hIcon);
                    _extensionIconCache[extension] = image;
                    return image;
                }
            }
            catch
            {
            }
            return null;
        }

        private async Task<ImageSource?> GetRecentItemIconAsync(RecentItem item)
        {
            // 尝试获取本地文件的系统自带缩略图/图标
            bool isLocalFile = false;
            string extension = string.Empty;

            if (!string.IsNullOrWhiteSpace(item.Location))
            {
                if (Uri.TryCreate(item.Location, UriKind.Absolute, out var uri))
                {
                    isLocalFile = uri.IsFile;
                    try
                    {
                        extension = System.IO.Path.GetExtension(uri.LocalPath);
                    }
                    catch { }
                }
                else
                {
                    isLocalFile = true;
                    try
                    {
                        extension = System.IO.Path.GetExtension(item.Location);
                    }
                    catch { }
                }
            }

            if (isLocalFile)
            {
                var systemIcon = await TryLoadFileIconAsync(item.Location);
                if (systemIcon != null)
                {
                    return systemIcon;
                }
            }

            // 对于云文档或者本地文件系统图标获取失败（文件可能不存在被移动）的情况，尝试根据后缀名生成假文件来获取系统类型图标
            if (!string.IsNullOrEmpty(extension))
            {
                var extIcon = await GetIconForExtensionAsync(extension);
                if (extIcon != null)
                {
                    return extIcon;
                }
            }

            // 如果连根据后缀获取图标也失败，触发兜底，使用应用主图标
            if (string.IsNullOrWhiteSpace(item.App)) return null;
            var baseName = item.App.ToLowerInvariant();

            if (_fileIconCache.TryGetValue(baseName, out var cached) && cached != null)
            {
                return cached;
            }

            var appIcon = await TryLoadFallbackIconAsync($"{baseName}.png");
            if (appIcon != null)
            {
                _fileIconCache[baseName] = appIcon;
                return appIcon;
            }

            return null;
        }

        private async Task RefreshRecentAsync(bool force = false)
        {
            if (!force && DateTimeOffset.Now - _lastRecentRefresh < RecentRefreshMinInterval)
            {
                return;
            }

            _lastRecentRefresh = DateTimeOffset.Now;
            var items = await Task.Run(() => Recent.GetItems(new RecentOptions
            {
                MaxItems = 50,
                UseInstalledAppsOnly = true,
                IncludeRegistry = true,
                IncludeCache = true,
                IncludeEdgeHistory = true
            }));

            // 预渲染并准备新项目列表
            var newItems = new System.Collections.Generic.List<RecentItem>();
            foreach (var item in items)
            {
                item.Icon = await GetRecentItemIconAsync(item);
                newItems.Add(item);
            }

            // 原地差异化更新 ObservableCollection 以消除 UI 闪烁
            for (int i = 0; i < newItems.Count; i++)
            {
                if (i < RecentItems.Count)
                {
                    // 仅当项目发生实质变动时才覆盖，避免无谓的 UI 重绘
                    if (RecentItems[i].Location != newItems[i].Location || 
                        RecentItems[i].TimestampText != newItems[i].TimestampText)
                    {
                        RecentItems[i] = newItems[i];
                    }
                }
                else
                {
                    RecentItems.Add(newItems[i]);
                }
            }

            // 移除多余的旧项目
            while (RecentItems.Count > newItems.Count)
            {
                RecentItems.RemoveAt(RecentItems.Count - 1);
            }
        }

        private void OfficeList_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (e.ClickedItem is not OfficeAppItem app)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(app.LaunchPath))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(app.LaunchPath) { UseShellExecute = true });
            }
            catch
            {
            }
        }

        private DateTime _lastOpenTime = DateTime.MinValue;

        private void RecentList_ItemClick(object sender, ItemClickEventArgs e)
        {
            if ((DateTime.Now - _lastOpenTime).TotalMilliseconds < 1000)
            {
                return;
            }
            _lastOpenTime = DateTime.Now;

            if (e.ClickedItem is not RecentItem recent)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(recent.Location))
            {
                return;
            }

            try
            {
                if (Uri.TryCreate(recent.Location, UriKind.Absolute, out var uri) && !uri.IsFile)
                {
                    Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
                    return;
                }

                if (File.Exists(recent.Location))
                {
                    Process.Start(new ProcessStartInfo(recent.Location) { UseShellExecute = true });
                }
            }
            catch
            {
            }
        }

        private void LoadWindowState()
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegistryPath);
                if (key != null)
                {
                    int width = (int)key.GetValue("WindowWidth", 800);
                    int height = (int)key.GetValue("WindowHeight", 600);
                    int x = (int)key.GetValue("WindowX", -1);
                    int y = (int)key.GetValue("WindowY", -1);
                    double leftColumnWidth = Convert.ToDouble(key.GetValue("LeftColumnWidth", 220.0));

                    if (width > 0 && height > 0)
                    {
                        this.AppWindow.Resize(new Windows.Graphics.SizeInt32(width, height));
                    }

                    if (x >= 0 && y >= 0)
                    {
                        var displayArea = Microsoft.UI.Windowing.DisplayArea.GetFromWindowId(
                            this.AppWindow.Id, Microsoft.UI.Windowing.DisplayAreaFallback.Primary);

                        var work = displayArea.WorkArea;
                        int workLeft = work.X;
                        int workTop = work.Y;
                        int workRight = work.X + work.Width;
                        int workBottom = work.Y + work.Height;

                        if (x < workRight && x + width > workLeft &&
                            y < workBottom && y + height > workTop)
                        {
                            this.AppWindow.Move(new Windows.Graphics.PointInt32(x, y));
                        }
                    }

                    // 恢复 GridSplitter 的分栏宽度
                    if (leftColumnWidth >= LeftColumn.MinWidth && leftColumnWidth <= LeftColumn.MaxWidth)
                    {
                        LeftColumn.Width = new GridLength(leftColumnWidth);
                    }
                }
            }
            catch { }
        }

        private void Window_Closed(object sender, WindowEventArgs args)
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegistryPath);
                if (key != null)
                {
                    var pos = this.AppWindow.Position;
                    var size = this.AppWindow.Size;

                    // 仅当窗口未最小化时保存坐标
                    if (size.Width > 0 && size.Height > 0)
                    {
                        key.SetValue("WindowX", pos.X, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("WindowY", pos.Y, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("WindowWidth", size.Width, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("WindowHeight", size.Height, Microsoft.Win32.RegistryValueKind.DWord);
                    }

                    // 保存左右栏宽度
                    key.SetValue("LeftColumnWidth", LeftColumn.Width.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }
            catch { }
        }

        private async void OnWindowActivated(object sender, WindowActivatedEventArgs args)
        {
            if (args.WindowActivationState != WindowActivationState.Deactivated)
            {
                await RefreshRecentAsync();
            }
        }

        private static ListViewItem? FindListViewItem(DependencyObject source)
        {
            var current = source;
            while (current != null)
            {
                if (current is ListViewItem item)
                {
                    return item;
                }

                current = VisualTreeHelper.GetParent(current);
            }

            return null;
        }

    }

    public sealed class OfficeAppItem : INotifyPropertyChanged
    {
        private ImageSource? _icon;

        public OfficeAppItem(string displayName, ImageSource? icon, string? launchPath)
        {
            DisplayName = displayName;
            _icon = icon;
            LaunchPath = launchPath;
        }

        public string DisplayName { get; }
        public string? LaunchPath { get; }
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
