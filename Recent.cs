using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace MsOfficeHub;

public sealed class RecentItem
{
	public RecentItem(string title, string location, DateTimeOffset timestamp, string app, string source, bool isCloud)
	{
		Title = title;
		Location = location;
		Timestamp = timestamp;
		App = app;
		Source = source;
		IsCloud = isCloud;
	}

	public string Title { get; }
	public string Location { get; }
	public DateTimeOffset Timestamp { get; }
	public string App { get; }
	public string Source { get; }
	public bool IsCloud { get; }

	public string DisplayLocation
	{
		get
		{
			if (IsCloud && Uri.TryCreate(Location, UriKind.Absolute, out var uri))
			{
				try
				{
					return Uri.UnescapeDataString(uri.AbsoluteUri);
				}
				catch
				{
					return Location;
				}
			}
			return Location;
		}
	}

	public string TimestampText => Timestamp.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
	public Microsoft.UI.Xaml.Media.ImageSource? Icon { get; set; }
}

public sealed class RecentOptions
{
	public int MaxItems { get; init; } = 50;
	public int MaxCacheItems { get; init; } = 200;
	public bool IncludeRegistry { get; init; } = true;
	public bool IncludeCache { get; init; } = true;
	public bool IncludeEdgeHistory { get; init; } = true;
	public bool UseInstalledAppsOnly { get; init; } = true;
	public int CacheMaxAgeDays { get; init; } = 90;
}

public static class Recent
{
	private static readonly RecentApp[] Apps =
	[
		new RecentApp("Word", new[] { "WINWORD.EXE" }),
		new RecentApp("Excel", new[] { "EXCEL.EXE" }),
		new RecentApp("PowerPoint", new[] { "POWERPNT.EXE" }),
		new RecentApp("Outlook", new[] { "OUTLOOK.EXE" }),
		new RecentApp("OneNote", new[] { "ONENOTE.EXE" }),
		new RecentApp("Access", new[] { "MSACCESS.EXE" }),
		new RecentApp("Publisher", new[] { "MSPUB.EXE" }),
		new RecentApp("Project", new[] { "WINPROJ.EXE" }),
		new RecentApp("Visio", new[] { "VISIO.EXE" })
	];

	public static IReadOnlyList<RecentItem> GetItems(RecentOptions? options = null)
	{
		options ??= new RecentOptions();

		var items = new List<RecentItem>();
		var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var appKeys = options.UseInstalledAppsOnly ? GetInstalledAppKeys() : Apps.Select(a => a.Key).ToHashSet(StringComparer.OrdinalIgnoreCase);

		if (options.IncludeRegistry)
		{
			AddRegistryItems(appKeys, items, seen);
		}

		if (options.IncludeCache)
		{
			AddCacheItems(items, seen, options.MaxCacheItems, options.CacheMaxAgeDays);
		}

		if (options.IncludeEdgeHistory)
		{
			AddEdgeHistoryItems(items, seen, options.CacheMaxAgeDays);
		}

		return items
			.OrderByDescending(item => item.Timestamp)
			.ThenBy(item => item.Title, StringComparer.OrdinalIgnoreCase)
			.Take(options.MaxItems)
			.ToList();
	}

	private static HashSet<string> GetInstalledAppKeys()
	{
		var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var detector = new OfficeDetector();

		foreach (var app in Apps)
		{
			var detection = detector.Detect(app.ExeNames, Array.Empty<string>(), ConfigService.ForceShowAllApps);
			if (detection.IsInstalled)
			{
				result.Add(app.Key);
			}
		}

		return result;
	}

	private static void AddRegistryItems(HashSet<string> appKeys, List<RecentItem> items, HashSet<string> seen)
	{
		foreach (var appKey in appKeys)
		{
			foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
			{
				using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
				var root = $"SOFTWARE\\Microsoft\\Office\\16.0\\{appKey}";
				using var fileMruKey = baseKey.OpenSubKey($"{root}\\File MRU");

				AddMruKeyItems(fileMruKey, appKey, "Registry", items, seen);
			}
		}
	}

	private static void AddMruKeyItems(RegistryKey? key, string appKey, string source, List<RecentItem> items, HashSet<string> seen)
	{
		if (key == null)
		{
			return;
		}

		var map = new Dictionary<int, string>();
		foreach (var name in key.GetValueNames())
		{
			if (name.Equals("MRUListEx", StringComparison.OrdinalIgnoreCase))
			{
				continue;
			}

			if (key.GetValue(name) is not string raw || string.IsNullOrWhiteSpace(raw))
			{
				continue;
			}

			if (!TryParseIndex(name, out var index))
			{
				continue;
			}

			map[index] = raw;
		}

		var order = ReadMruListEx(key);
		IEnumerable<int> orderedIndexes = order.Count > 0 ? order : map.Keys.OrderBy(i => i);

		foreach (var index in orderedIndexes)
		{
			if (!map.TryGetValue(index, out var raw))
			{
				continue;
			}

			var location = ExtractLocation(raw);
			if (string.IsNullOrWhiteSpace(location))
			{
				continue;
			}

			var normalized = NormalizeLocation(location);
			if (!seen.Add(normalized))
			{
				continue;
			}

			var timestamp = TryParseTimestamp(raw) ?? GetFileTimestamp(location) ?? DateTimeOffset.MinValue;
			var title = BuildTitle(location);
			var isCloud = IsCloudLocation(location);

			if (!isCloud)
			{
				try
				{
					var checkLocation = location;
					if (Uri.TryCreate(location, UriKind.Absolute, out var uri) && uri.IsFile)
					{
						checkLocation = uri.LocalPath;
					}
					
					if (!System.IO.File.Exists(checkLocation))
					{
						continue;
					}
				}
				catch { }
			}

			items.Add(new RecentItem(title, location, timestamp, appKey, source, isCloud));
		}
	}

	private static string? GuessAppFromExtension(string path)
	{
		try
		{
			var target = path;
			if (Uri.TryCreate(path, UriKind.Absolute, out var uri))
			{
				target = uri.LocalPath;
			}

			var ext = Path.GetExtension(target)?.ToLowerInvariant();
			if (string.IsNullOrEmpty(ext)) return null;
			if (ext.StartsWith(".doc") || ext == ".rtf" || ext == ".txt" || ext == ".odt") return "Word";
			if (ext.StartsWith(".xls") || ext == ".csv" || ext == ".tsv") return "Excel";
			if (ext.StartsWith(".ppt") || ext == ".pps") return "PowerPoint";
			if (ext.StartsWith(".vsd") || ext == ".vdx") return "Visio";
			if (ext.StartsWith(".pub")) return "Publisher";
			if (ext.StartsWith(".mpp")) return "Project";
			if (ext.StartsWith(".one")) return "OneNote";
			if (ext.StartsWith(".accdb") || ext == ".mdb") return "Access";
		}
		catch
		{
		}
		return null;
	}

	private static void AddCacheItems(List<RecentItem> items, HashSet<string> seen, int maxItems, int maxAgeDays)
	{
		var candidates = new List<string>();
		var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
		var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

		var dirs = new[]
		{
			Path.Combine(appData, "Microsoft", "Office", "Recent"),
			Path.Combine(appData, "Microsoft", "Office", "Recent", "Documents"),
			Path.Combine(appData, "Microsoft", "Office", "Recent", "Templates"),
			Path.Combine(localAppData, "Microsoft", "Office", "Recent")
		};

		foreach (var dir in dirs)
		{
			if (!Directory.Exists(dir))
			{
				continue;
			}

			try
			{
				candidates.AddRange(Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories));
			}
			catch
			{
			}
		}

		var cutoff = DateTimeOffset.Now.AddDays(-Math.Max(1, maxAgeDays));
		foreach (var path in candidates
					 .Where(p => p.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase) || p.EndsWith(".url", StringComparison.OrdinalIgnoreCase))
					 .OrderByDescending(GetFileTimestamp)
					 .Take(maxItems))
		{
			var target = ResolveCacheTarget(path);
			if (string.IsNullOrWhiteSpace(target))
			{
				continue;
			}

			var normalized = NormalizeLocation(target);
			if (!seen.Add(normalized))
			{
				continue;
			}

			var isCloud = IsCloudLocation(target);
			if (!isCloud && Directory.Exists(target))
			{
				continue;
			}

			// 防御性过滤云端文件夹：无拓展名或以斜杠结尾大概率是文件夹
			var guessedApp = GuessAppFromExtension(target);
			if (guessedApp == null && isCloud)
			{
				if (target.EndsWith("/") || target.EndsWith("\\") || !Path.HasExtension(target))
				{
					continue;
				}
			}

			var timestamp = GetFileTimestamp(path) ?? DateTimeOffset.MinValue;
			if (timestamp < cutoff)
			{
				continue;
			}

			var finalApp = guessedApp ?? "Word"; // 默认兜底分配给Word或公用文档图标
			var title = BuildTitle(target);
			items.Add(new RecentItem(title, target, timestamp, finalApp, "Cache", isCloud));
		}
	}

	private static void AddEdgeHistoryItems(List<RecentItem> items, HashSet<string> seen, int maxAgeDays)
	{
		try
		{
			var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
			var edgeHistoryPath = Path.Combine(localAppData, "Microsoft", "Edge", "User Data", "Default", "History");

			if (!File.Exists(edgeHistoryPath))
			{
				return;
			}

			var tempFile = Path.Combine(Path.GetTempPath(), $"EdgeHistory_{Guid.NewGuid():N}.db");
			try
			{
				File.Copy(edgeHistoryPath, tempFile, true);
				var cutoffCut = DateTimeOffset.Now.AddDays(-maxAgeDays);
				var cutoffMicroseconds = cutoffCut.ToFileTime() / 10;

				using var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={tempFile}");
				connection.Open();

				using var command = connection.CreateCommand();
				// 筛选本地和网络 PDF
				command.CommandText = @"
					SELECT url, title, last_visit_time 
					FROM urls 
					WHERE url LIKE 'file://%.pdf' OR url LIKE '%.pdf' OR url LIKE '%.pdf?%'
					ORDER BY last_visit_time DESC 
					LIMIT 100";

				using var reader = command.ExecuteReader();
				while (reader.Read())
				{
					var url = reader.GetString(0);
					var rawTitle = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
					var lastVisitMicros = reader.GetInt64(2);

					if (lastVisitMicros < cutoffMicroseconds)
					{
						continue; // 过旧的记录
					}

					var target = url;
					bool isCloud = false;
					if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
					{
						if (uri.IsFile)
						{
							target = uri.LocalPath;
						}
						else
						{
							isCloud = true;
						}
					}

					var normalized = NormalizeLocation(target);
					if (!seen.Add(normalized))
					{
						continue;
					}

					if (!isCloud)
					{
						if (!File.Exists(target))
						{
							continue;
						}
					}

					var timestamp = DateTimeOffset.FromFileTime(lastVisitMicros * 10);
					var title = string.IsNullOrWhiteSpace(rawTitle) ? BuildTitle(target) : rawTitle;

					items.Add(new RecentItem(title, target, timestamp, "Edge", "Browser", isCloud));
				}
			}
			finally
			{
				if (File.Exists(tempFile))
				{
					try { File.Delete(tempFile); } catch { }
				}
			}
		}
		catch
		{
		}
	}

	private static List<int> ReadMruListEx(RegistryKey key)
	{
		var list = new List<int>();
		if (key.GetValue("MRUListEx") is not byte[] raw || raw.Length < 4)
		{
			return list;
		}

		for (var i = 0; i + 3 < raw.Length; i += 4)
		{
			var value = BitConverter.ToInt32(raw, i);
			if (value == -1)
			{
				break;
			}

			list.Add(value);
		}

		return list;
	}

	private static bool TryParseIndex(string name, out int index)
	{
		index = -1;
		if (name.StartsWith("Item ", StringComparison.OrdinalIgnoreCase))
		{
			return int.TryParse(name.Substring(5), NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
		}

		return int.TryParse(name, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
	}

	private static string? ExtractLocation(string raw)
	{
		if (string.IsNullOrWhiteSpace(raw))
		{
			return null;
		}

		var star = raw.LastIndexOf('*');
		var part = star >= 0 ? raw[(star + 1)..] : raw;
		part = part.Trim();
		if (part.Length == 0)
		{
			return null;
		}

		if (part.StartsWith("Z", StringComparison.OrdinalIgnoreCase))
		{
			var trimmed = part[1..];
			if (trimmed.StartsWith("\\\\", StringComparison.OrdinalIgnoreCase) || trimmed.StartsWith("http", StringComparison.OrdinalIgnoreCase) || (trimmed.Length > 1 && trimmed[1] == ':'))
			{
				part = trimmed;
			}
		}

		if (part.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
		{
			if (Uri.TryCreate(part, UriKind.Absolute, out var uri) && uri.IsFile)
			{
				return uri.LocalPath;
			}
		}

		return part;
	}

	private static DateTimeOffset? TryParseTimestamp(string raw)
	{
		if (string.IsNullOrWhiteSpace(raw))
		{
			return null;
		}

		var star = raw.IndexOf('*');
		if (raw.Length > 2 && raw[0] == 'T' && star > 1)
		{
			var token = raw.Substring(1, star - 1);
			if (token.Length >= 16 && long.TryParse(token[..16], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var fileTime))
			{
				try
				{
					return DateTimeOffset.FromFileTime(fileTime);
				}
				catch
				{
					return null;
				}
			}
		}

		return null;
	}

	private static DateTimeOffset? GetFileTimestamp(string? path)
	{
		try
		{
			if (string.IsNullOrWhiteSpace(path))
			{
				return null;
			}

			if (File.Exists(path))
			{
				return new DateTimeOffset(File.GetLastWriteTimeUtc(path));
			}
		}
		catch
		{
			return null;
		}

		return null;
	}

	private static string NormalizeLocation(string location)
	{
		if (string.IsNullOrWhiteSpace(location))
		{
			return string.Empty;
		}

		if (Uri.TryCreate(location, UriKind.Absolute, out var uri) && uri.IsAbsoluteUri)
		{
			return uri.GetComponents(UriComponents.AbsoluteUri, UriFormat.SafeUnescaped).Trim().ToLowerInvariant();
		}

		var normalized = location.Trim();
		normalized = normalized.Replace('/', '\\');
		return normalized.ToLowerInvariant();
	}

	private static string BuildTitle(string location)
	{
		if (Uri.TryCreate(location, UriKind.Absolute, out var uri) && uri.IsAbsoluteUri && !uri.IsFile)
		{
			try
			{
				var unescaped = Uri.UnescapeDataString(uri.AbsolutePath);
				var fileName = Path.GetFileName(unescaped);
				if (!string.IsNullOrWhiteSpace(fileName))
				{
					return fileName;
				}
				return unescaped.Trim('/');
			}
			catch
			{
				return uri.AbsolutePath.Trim('/');
			}
		}

		try
		{
			return Path.GetFileName(location);
		}
		catch
		{
			return location;
		}
	}

	private static bool IsCloudLocation(string location)
	{
		if (!Uri.TryCreate(location, UriKind.Absolute, out var uri))
		{
			return false;
		}

		if (uri.IsFile)
		{
			return false;
		}

		return uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) || uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase);
	}

	private static string? ResolveCacheTarget(string path)
	{
		if (path.EndsWith(".url", StringComparison.OrdinalIgnoreCase))
		{
			return ResolveUrlShortcut(path);
		}

		if (path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
		{
			return ResolveShellShortcut(path);
		}

		return null;
	}

	private static string? ResolveUrlShortcut(string path)
	{
		try
		{
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			var encoding = Encoding.GetEncoding(0);

			foreach (var line in File.ReadAllLines(path, encoding))
			{
				if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
				{
					return line.Substring(4).Trim();
				}
			}
		}
		catch
		{
			return null;
		}

		return null;
	}

	private static string? ResolveShellShortcut(string path)
	{
		try
		{
			var link = (IShellLinkW)new ShellLink();
			((IPersistFile)link).Load(path, 0);
			var builder = new StringBuilder(260);
			link.GetPath(builder, builder.Capacity, out _, 0);
			var result = builder.ToString();
			return string.IsNullOrWhiteSpace(result) ? null : result;
		}
		catch
		{
			return null;
		}
	}

	private sealed class RecentApp
	{
		public RecentApp(string key, string[] exeNames)
		{
			Key = key;
			ExeNames = exeNames;
		}

		public string Key { get; }
		public string[] ExeNames { get; }
	}

	[ComImport]
	[Guid("00021401-0000-0000-C000-000000000046")]
	private class ShellLink
	{
	}

	[ComImport]
	[Guid("000214F9-0000-0000-C000-000000000046")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	private interface IShellLinkW
	{
		void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, out WIN32_FIND_DATAW pfd, uint fFlags);
		void GetIDList(out IntPtr ppidl);
		void SetIDList(IntPtr pidl);
		void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
		void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
		void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
		void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
		void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
		void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
		void GetHotkey(out short wHotkey);
		void SetHotkey(short wHotkey);
		void GetShowCmd(out int iShowCmd);
		void SetShowCmd(int iShowCmd);
		void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int iIcon);
		void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
		void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
		void Resolve(IntPtr hwnd, uint fFlags);
		void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	private struct WIN32_FIND_DATAW
	{
		public uint dwFileAttributes;
		public FILETIME ftCreationTime;
		public FILETIME ftLastAccessTime;
		public FILETIME ftLastWriteTime;
		public uint nFileSizeHigh;
		public uint nFileSizeLow;
		public uint dwReserved0;
		public uint dwReserved1;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
		public string cFileName;
		[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
		public string cAlternateFileName;
	}

	[ComImport]
	[Guid("0000010b-0000-0000-C000-000000000046")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	private interface IPersistFile
	{
		void GetClassID(out Guid pClassID);
		[PreserveSig]
		int IsDirty();
		void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
		void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);
		void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
		void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
	}
}
