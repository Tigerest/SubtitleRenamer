using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace SubtitleRenamer;

internal sealed class OnlineSubtitleCache : IDisposable
{
    private readonly Dictionary<string, AcgripSearchSnapshot> _searches = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _downloads = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _knownSubtitlePaths = new(StringComparer.OrdinalIgnoreCase);
    private bool _disposed;

    public OnlineSubtitleCache()
    {
        var parent = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SubtitleRenamer",
            "AcgripCache");

        Directory.CreateDirectory(parent);
        RootDirectory = Path.Combine(parent, $"run-{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}"[..31]);
        Directory.CreateDirectory(RootDirectory);
        Directory.CreateDirectory(DownloadsDirectory);
        Directory.CreateDirectory(ExtractsDirectory);
    }

    public string RootDirectory { get; }
    public string DownloadsDirectory => Path.Combine(RootDirectory, "downloads");
    public string ExtractsDirectory => Path.Combine(RootDirectory, "extracts");
    public List<OnlineSubtitleItem> Subtitles { get; } = new();

    public bool TryGetSearch(string query, out AcgripSearchSnapshot snapshot)
    {
        return _searches.TryGetValue(NormalizeQueryKey(query), out snapshot!);
    }

    public void SaveSearch(AcgripSearchSnapshot snapshot)
    {
        _searches[NormalizeQueryKey(snapshot.Query)] = snapshot;
        SaveManifest();
    }

    public bool TryGetDownloadedAttachment(AcgripAttachment attachment, out string path)
    {
        return _downloads.TryGetValue(attachment.CacheKey, out path!) && File.Exists(path);
    }

    public string GetDownloadPath(AcgripAttachment attachment, string fileName)
    {
        var directory = Path.Combine(DownloadsDirectory, SafeName(attachment.CacheKey));
        Directory.CreateDirectory(directory);
        return Path.Combine(directory, SafeFileName(fileName));
    }

    public void SaveDownloadedAttachment(AcgripAttachment attachment, string path)
    {
        _downloads[attachment.CacheKey] = path;
        SaveManifest();
    }

    public string GetExtractDirectory(AcgripAttachment attachment)
    {
        var directory = Path.Combine(ExtractsDirectory, SafeName(attachment.CacheKey));
        Directory.CreateDirectory(directory);
        return directory;
    }

    public int AddExtractedSubtitles(IEnumerable<string> paths, string extractDirectory, AcgripThreadFloor floor, AcgripAttachment attachment)
    {
        var added = 0;
        var extractRoot = Path.GetFullPath(extractDirectory);
        foreach (var path in paths.Select(Path.GetFullPath))
        {
            if (!_knownSubtitlePaths.Add(path))
            {
                continue;
            }

            Subtitles.Add(new OnlineSubtitleItem(path, floor.Thread.Title, floor.FloorLabel, attachment.Name, GetFolderLabel(extractRoot, path)));
            added++;
        }

        if (added > 0)
        {
            SaveManifest();
        }

        return added;
    }

    public int RemoveSubtitles(IEnumerable<string> paths, bool deleteFiles)
    {
        var pathSet = paths.Select(Path.GetFullPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var removed = Subtitles.RemoveAll(item => pathSet.Contains(Path.GetFullPath(item.FullPath)));
        foreach (var path in pathSet)
        {
            _knownSubtitlePaths.Remove(path);
            if (!deleteFiles)
            {
                continue;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                    DeleteEmptyParents(Path.GetDirectoryName(path));
                }
            }
            catch
            {
                // The UI list should still be cleaned up if a cached file is locked.
            }
        }

        if (removed > 0)
        {
            SaveManifest();
        }

        return removed;
    }

    public void Save()
    {
        SaveManifest();
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        try
        {
            if (Directory.Exists(RootDirectory))
            {
                Directory.Delete(RootDirectory, recursive: true);
            }
        }
        catch
        {
            // Best-effort cleanup. The cache is intentionally per run and non-critical.
        }
    }

    private void SaveManifest()
    {
        try
        {
            var manifest = new
            {
                CreatedAt = DateTime.Now,
                Searches = _searches.Values.Select(item => new { item.Query, item.CreatedAt, FloorCount = item.Floors.Count }),
                Downloads = _downloads,
                Subtitles = Subtitles.Select(item => new
                {
                    item.FullPath,
                    item.ThreadTitle,
                    item.FloorLabel,
                    item.AttachmentName,
                    item.FolderLabel,
                    item.Imported
                })
            };
            var json = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(RootDirectory, "manifest.json"), json, Encoding.UTF8);
        }
        catch
        {
            // A broken manifest must not break subtitle importing.
        }
    }

    private static string NormalizeQueryKey(string query)
    {
        return query.Trim().ToLowerInvariant();
    }

    private static string SafeName(string text)
    {
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)))[..16].ToLowerInvariant();
        return hash;
    }

    private static string SafeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(fileName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return cleaned.Length > 0 ? cleaned : "attachment.bin";
    }

    private static string GetFolderLabel(string extractRoot, string subtitlePath)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(subtitlePath)) ?? extractRoot;
        var relative = Path.GetRelativePath(extractRoot, directory);
        return string.IsNullOrWhiteSpace(relative) || relative == "."
            ? "根目录"
            : relative.Replace(Path.DirectorySeparatorChar, '/').Replace(Path.AltDirectorySeparatorChar, '/');
    }

    private static void DeleteEmptyParents(string? directory)
    {
        while (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
        {
            if (Directory.EnumerateFileSystemEntries(directory).Any())
            {
                return;
            }

            var parent = Directory.GetParent(directory)?.FullName;
            Directory.Delete(directory);
            directory = parent;
        }
    }
}
