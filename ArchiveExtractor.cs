using SharpCompress.Common;
using SharpCompress.Readers;

namespace SubtitleRenamer;

internal static class ArchiveExtractor
{
    private static readonly HashSet<string> SubtitleExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".ass", ".ssa", ".srt", ".sub", ".idx", ".sup", ".vtt", ".smi", ".ttml", ".dfxp"
    };

    private static readonly HashSet<string> ArchiveExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".zip", ".7z", ".rar"
    };

    public static bool IsSupportedAttachmentName(string name)
    {
        var extension = Path.GetExtension(name);
        return SubtitleExtensions.Contains(extension) || ArchiveExtensions.Contains(extension);
    }

    public static bool IsSubtitle(string path)
    {
        return SubtitleExtensions.Contains(Path.GetExtension(path));
    }

    public static async Task<IReadOnlyList<string>> ExtractSubtitlesAsync(string sourcePath, string destinationDirectory, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(destinationDirectory);

        if (IsSubtitle(sourcePath))
        {
            var target = GetUniquePath(destinationDirectory, Path.GetFileName(sourcePath));
            File.Copy(sourcePath, target, overwrite: false);
            return new[] { target };
        }

        if (!ArchiveExtensions.Contains(Path.GetExtension(sourcePath)))
        {
            return Array.Empty<string>();
        }

        return await Task.Run(() =>
        {
            var extracted = new List<string>();
            using var reader = ReaderFactory.OpenReader(sourcePath, new ReaderOptions());
            while (reader.MoveToNextEntry())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var entry = reader.Entry;
                if (entry.IsDirectory)
                {
                    continue;
                }

                var fileName = Path.GetFileName(entry.Key);
                if (string.IsNullOrWhiteSpace(fileName) || !IsSubtitle(fileName))
                {
                    continue;
                }

                var target = GetUniquePath(destinationDirectory, BuildSafeRelativePath(entry.Key ?? fileName));
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                using (var output = File.Create(target))
                {
                    reader.WriteEntryTo(output);
                }

                extracted.Add(target);
            }

            return (IReadOnlyList<string>)extracted;
        }, cancellationToken);
    }

    private static string GetUniquePath(string directory, string relativePath)
    {
        var safeName = BuildSafeRelativePath(relativePath);
        var candidate = Path.Combine(directory, safeName);
        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var parent = Path.GetDirectoryName(candidate) ?? directory;
        var stem = Path.GetFileNameWithoutExtension(candidate);
        var extension = Path.GetExtension(candidate);
        for (var i = 2; ; i++)
        {
            candidate = Path.Combine(parent, $"{Path.GetFileName(stem)}-{i}{extension}");
            if (!File.Exists(candidate))
            {
                return candidate;
            }
        }
    }

    private static string BuildSafeRelativePath(string path)
    {
        var normalized = path.Replace('\\', '/');
        var parts = normalized
            .Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(part => part != "." && part != "..")
            .Select(SafeFileName)
            .Where(part => part.Length > 0)
            .ToArray();

        return parts.Length == 0 ? "subtitle" : Path.Combine(parts);
    }

    private static string SafeFileName(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(value.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return cleaned.Length > 0 ? cleaned : "_";
    }
}
