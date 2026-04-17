using System.Diagnostics;
using SharpCompress.Archives;
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

        try
        {
            return await ExtractWithSharpCompressAsync(sourcePath, destinationDirectory, cancellationToken);
        }
        catch (Exception) when (!cancellationToken.IsCancellationRequested
            && Path.GetExtension(sourcePath).Equals(".7z", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                return await ExtractWithSharpCompressSevenZipArchiveAsync(sourcePath, destinationDirectory, cancellationToken);
            }
            catch (Exception archiveException) when (!cancellationToken.IsCancellationRequested)
            {
                return await ExtractWithSevenZipAsync(sourcePath, destinationDirectory, archiveException, cancellationToken);
            }
        }
    }

    private static async Task<IReadOnlyList<string>> ExtractWithSharpCompressAsync(string sourcePath, string destinationDirectory, CancellationToken cancellationToken)
    {
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

    private static async Task<IReadOnlyList<string>> ExtractWithSharpCompressSevenZipArchiveAsync(
        string sourcePath,
        string destinationDirectory,
        CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            var extracted = new List<string>();
            using var archive = ArchiveFactory.OpenArchive(sourcePath, new ReaderOptions());
            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
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
                using var input = entry.OpenEntryStream();
                using var output = File.Create(target);
                input.CopyTo(output);
                extracted.Add(target);
            }

            return (IReadOnlyList<string>)extracted;
        }, cancellationToken);
    }

    private static async Task<IReadOnlyList<string>> ExtractWithSevenZipAsync(
        string sourcePath,
        string destinationDirectory,
        Exception originalException,
        CancellationToken cancellationToken)
    {
        var sevenZipPath = FindSevenZipExecutable();
        if (sevenZipPath is null)
        {
            throw new InvalidOperationException(
                $"内置解压器无法打开这个 7z 压缩包，并且没有找到本机 7-Zip。请安装 7-Zip 后重试。原始错误：{originalException.Message}",
                originalException);
        }

        var tempDirectory = Path.Combine(destinationDirectory, $"_7z_extract_{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempDirectory);

        try
        {
            var outputDirectoryArg = $"-o{tempDirectory}";
            var startInfo = new ProcessStartInfo
            {
                FileName = sevenZipPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            startInfo.ArgumentList.Add("x");
            startInfo.ArgumentList.Add("-y");
            startInfo.ArgumentList.Add(outputDirectoryArg);
            startInfo.ArgumentList.Add(sourcePath);

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("无法启动 7-Zip 解压进程。");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync(cancellationToken);
            var stdout = await stdoutTask;
            var stderr = await stderrTask;

            if (process.ExitCode != 0)
            {
                var detail = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
                throw new InvalidOperationException($"7-Zip 解压失败：{detail.Trim()}");
            }

            return await Task.Run(() =>
            {
                var extracted = new List<string>();
                foreach (var subtitlePath in Directory.EnumerateFiles(tempDirectory, "*", SearchOption.AllDirectories)
                    .Where(IsSubtitle))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var relative = Path.GetRelativePath(tempDirectory, subtitlePath);
                    var target = GetUniquePath(destinationDirectory, relative);
                    Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                    File.Copy(subtitlePath, target, overwrite: false);
                    extracted.Add(target);
                }

                return (IReadOnlyList<string>)extracted;
            }, cancellationToken);
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, recursive: true);
                }
            }
            catch
            {
                // A failed cleanup should not hide a successful extraction.
            }
        }
    }

    private static string? FindSevenZipExecutable()
    {
        foreach (var executable in new[] { "7z.exe", "7za.exe", "7zr.exe" })
        {
            var path = FindOnPath(executable);
            if (path is not null)
            {
                return path;
            }
        }

        foreach (var path in GetCommonSevenZipPaths())
        {
            if (File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    private static string? FindOnPath(string executable)
    {
        var pathVariable = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(pathVariable))
        {
            return null;
        }

        foreach (var directory in pathVariable.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            try
            {
                var candidate = Path.Combine(directory, executable);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch
            {
            }
        }

        return null;
    }

    private static IEnumerable<string> GetCommonSevenZipPaths()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (!string.IsNullOrWhiteSpace(programFiles))
        {
            yield return Path.Combine(programFiles, "7-Zip", "7z.exe");
        }

        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        if (!string.IsNullOrWhiteSpace(programFilesX86))
        {
            yield return Path.Combine(programFilesX86, "7-Zip", "7z.exe");
        }
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
