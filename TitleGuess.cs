using System.Text.RegularExpressions;

namespace SubtitleRenamer;

internal static class TitleGuess
{
    private static readonly HashSet<string> TechnicalTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "1080p", "720p", "2160p", "480p", "4k", "8k", "bdrip", "bdrips",
        "bluray", "blu-ray", "web-dl", "webrip", "web", "tv", "dvdrip",
        "hevc", "avc", "h264", "h265", "x264", "x265", "10bit", "8bit",
        "hi10p", "ma10p", "aac", "flac", "dts", "opus", "mp3", "hdr", "sdr",
        "sub", "subs", "multi", "dual", "chs", "cht", "sc", "tc", "gb", "big5"
    };

    public static string GuessFromVideoName(string fileName)
    {
        var stem = Path.GetFileNameWithoutExtension(fileName);
        if (string.IsNullOrWhiteSpace(stem))
        {
            return "";
        }

        var text = Regex.Replace(stem, @"^\s*(?:[\[\(【（][^\]\)】）]{1,40}[\]\)】）]\s*)+", "");
        text = Regex.Replace(text, @"[\._]+", " ");
        text = Regex.Replace(text, @"(?i)\bS\d{1,2}E\d{1,3}\b.*$", "");
        text = Regex.Replace(text, @"(?i)\b\d{1,2}x\d{1,3}\b.*$", "");
        text = Regex.Replace(text, @"\s-\s\d{1,3}(?:v\d+)?\b.*$", "");
        text = Regex.Replace(text, @"(?i)\bEP?\s*\d{1,3}(?:v\d+)?\b.*$", "");
        text = Regex.Replace(text, @"第\s*\d{1,3}\s*[话話集].*$", "");

        text = RemoveTrailingBracketTags(text);
        var tokens = text
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(token => !TechnicalTokens.Contains(token))
            .ToList();

        var guessed = string.Join(" ", tokens).Trim(' ', '-', '_', '.', '　');
        return guessed.Length > 0 ? guessed : stem;
    }

    private static string RemoveTrailingBracketTags(string text)
    {
        while (true)
        {
            var trimmed = text.Trim();
            var match = Regex.Match(trimmed, @"(?:[\[\(【（](?<tag>[^\]\)】）]+)[\]\)】）])\s*$");
            if (!match.Success)
            {
                return trimmed;
            }

            var tag = match.Groups["tag"].Value.Trim();
            var tagTokens = tag.Split(new[] { ' ', '_', '-', '.', ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (tagTokens.Length == 0 || tagTokens.Any(token => !TechnicalTokens.Contains(token)))
            {
                return trimmed;
            }

            text = trimmed[..match.Index];
        }
    }
}
