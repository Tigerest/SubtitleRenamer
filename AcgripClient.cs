using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using HapHtmlDocument = HtmlAgilityPack.HtmlDocument;
using HapHtmlEntity = HtmlAgilityPack.HtmlEntity;

namespace SubtitleRenamer;

internal sealed class AcgripClient : IDisposable
{
    private static readonly Uri BaseUri = new("https://bbs.acgrip.com/");
    private static readonly TimeSpan MinimumRequestInterval = TimeSpan.FromSeconds(2);
    private static readonly HashSet<string> PreferredForums = new(StringComparer.OrdinalIgnoreCase)
    {
        "ACG字幕分享",
        "旧物/老番区"
    };

    private static readonly HashSet<string> DeprioritizedForums = new(StringComparer.OrdinalIgnoreCase)
    {
        "灌水聊天",
        "作品报错区",
        "推荐/开坑请愿",
        "积分合作讨论"
    };

    private readonly HttpClient _http;
    private readonly SemaphoreSlim _requestLock = new(1, 1);
    private DateTimeOffset _lastRequest = DateTimeOffset.MinValue;

    public AcgripClient()
    {
        var handler = new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.All,
            UseCookies = true,
            CookieContainer = new CookieContainer()
        };

        _http = new HttpClient(handler)
        {
            BaseAddress = BaseUri,
            Timeout = TimeSpan.FromSeconds(60)
        };
        _http.DefaultRequestHeaders.UserAgent.ParseAdd(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
            "(KHTML, like Gecko) Chrome/135.0 Safari/537.36 SubtitleRenamer/1.0");
        _http.DefaultRequestHeaders.Accept.ParseAdd("text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
        _http.DefaultRequestHeaders.AcceptLanguage.ParseAdd("zh-CN,zh;q=0.9,en;q=0.6");
    }

    public async Task<IReadOnlyList<AcgripThreadResult>> SearchAsync(string query, int maxResults, CancellationToken cancellationToken)
    {
        var url = $"search.php?mod=forum&srchtxt={Uri.EscapeDataString(query)}&searchsubmit=yes";
        var html = await GetStringAsync(url, referer: null, cancellationToken);
        var doc = LoadDocument(html);
        var items = doc.DocumentNode.SelectNodes("//*[@id='threadlist']//li[contains(concat(' ', normalize-space(@class), ' '), ' pbw ')]")
            ?? Enumerable.Empty<HtmlNode>();

        var results = new Dictionary<int, AcgripThreadResult>();
        foreach (var item in items)
        {
            var link = item.SelectSingleNode(".//h3//a[@href]") ?? item.SelectSingleNode(".//a[contains(@href,'tid=') or contains(@href,'thread-')]");
            if (link is null)
            {
                continue;
            }

            var href = ToAbsoluteUrl(link.GetAttributeValue("href", ""));
            var tid = ParseTid(href);
            if (tid <= 0)
            {
                continue;
            }

            var forumName = CleanText(item.SelectSingleNode(".//p//span//a[contains(concat(' ', normalize-space(@class), ' '), ' xi1 ')]")?.InnerText ?? "");
            var paragraphs = (item.SelectNodes(".//p") ?? Enumerable.Empty<HtmlNode>()).ToList();
            var snippet = paragraphs.Count > 1 ? CleanText(paragraphs[1].InnerText) : "";
            var title = CleanText(link.InnerText);
            var score = ScoreSearchResult(title, snippet, forumName, query);
            results[tid] = new AcgripThreadResult(tid, title, href, forumName, snippet, score);
        }

        return results.Values
            .OrderByDescending(item => item.Score)
            .ThenBy(item => item.Title)
            .Take(maxResults)
            .ToList();
    }

    public async Task<IReadOnlyList<AcgripThreadFloor>> FetchSubtitleFloorsAsync(
        AcgripThreadResult thread,
        int maxPages,
        CancellationToken cancellationToken)
    {
        var floors = new List<AcgripThreadFloor>();
        var seenPostIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenAttachmentUrls = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var floorNumber = 0;

        for (var page = 1; page <= maxPages; page++)
        {
            var url = page == 1
                ? $"forum.php?mod=viewthread&tid={thread.Tid}"
                : $"forum.php?mod=viewthread&tid={thread.Tid}&page={page}";
            var html = await GetStringAsync(url, thread.Url, cancellationToken);
            var doc = LoadDocument(html);
            var links = doc.DocumentNode.SelectNodes("//a[contains(@href,'mod=attachment') and @href]")
                ?? Enumerable.Empty<HtmlNode>();
            var grouped = new Dictionary<string, List<HtmlNode>>(StringComparer.OrdinalIgnoreCase);

            foreach (var link in links)
            {
                var post = FindPostNode(link);
                var postId = post?.GetAttributeValue("id", "") ?? $"page-{page}-unknown";
                if (!grouped.TryGetValue(postId, out var postLinks))
                {
                    postLinks = new List<HtmlNode>();
                    grouped[postId] = postLinks;
                }

                postLinks.Add(link);
            }

            foreach (var (postId, postLinks) in grouped)
            {
                if (!seenPostIds.Add(postId))
                {
                    continue;
                }

                var attachments = new List<AcgripAttachment>();
                foreach (var link in postLinks)
                {
                    var href = ToAbsoluteUrl(link.GetAttributeValue("href", ""));
                    if (!seenAttachmentUrls.Add(href))
                    {
                        continue;
                    }

                    var name = InferAttachmentName(link);
                    if (!ArchiveExtractor.IsSupportedAttachmentName(name))
                    {
                        continue;
                    }

                    attachments.Add(new AcgripAttachment(
                        name,
                        href,
                        Path.GetExtension(name).ToLowerInvariant(),
                        BuildAttachmentKey(href)));
                }

                if (attachments.Count == 0)
                {
                    continue;
                }

                floorNumber++;
                var post = postLinks.Select(FindPostNode).FirstOrDefault(node => node is not null);
                floors.Add(new AcgripThreadFloor(
                    thread,
                    page,
                    floorNumber,
                    ExtractAuthor(post),
                    ExtractPostedAt(post),
                    ExtractExcerpt(post),
                    attachments));
            }
        }

        return floors;
    }

    public async Task<string> DownloadAttachmentAsync(
        AcgripAttachment attachment,
        string targetPath,
        string referer,
        IProgress<AcgripDownloadProgress>? progress,
        CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, attachment.Url);
        request.Headers.Referrer = new Uri(referer);
        request.Headers.Accept.Clear();
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
        using var response = await SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();

        var dispositionName = ParseDispositionFileName(response.Content.Headers.ContentDisposition?.ToString() ?? "");
        if (!string.IsNullOrWhiteSpace(dispositionName))
        {
            targetPath = Path.Combine(Path.GetDirectoryName(targetPath)!, SafeFileName(dispositionName));
        }

        var totalBytes = response.Content.Headers.ContentLength;
        var startedAt = DateTimeOffset.UtcNow;
        var received = 0L;
        var buffer = new byte[1024 * 64];
        await using var input = await response.Content.ReadAsStreamAsync(cancellationToken);
        await using var output = File.Create(targetPath);
        while (true)
        {
            var read = await input.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken);
            if (read == 0)
            {
                break;
            }

            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
            received += read;
            var elapsed = Math.Max(0.1, (DateTimeOffset.UtcNow - startedAt).TotalSeconds);
            progress?.Report(new AcgripDownloadProgress(attachment.Name, received, totalBytes, received / elapsed));
        }

        return targetPath;
    }

    public void Dispose()
    {
        _http.Dispose();
        _requestLock.Dispose();
    }

    private async Task<string> GetStringAsync(string url, string? referer, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, url);
        if (!string.IsNullOrWhiteSpace(referer))
        {
            request.Headers.Referrer = new Uri(referer);
        }

        using var response = await SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();
        var bytes = await response.Content.ReadAsByteArrayAsync(cancellationToken);
        return Encoding.UTF8.GetString(bytes);
    }

    private async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        await _requestLock.WaitAsync(cancellationToken);
        try
        {
            var wait = _lastRequest + MinimumRequestInterval - DateTimeOffset.UtcNow;
            if (wait > TimeSpan.Zero)
            {
                await Task.Delay(wait, cancellationToken);
            }

            var response = await _http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            _lastRequest = DateTimeOffset.UtcNow;
            return response;
        }
        finally
        {
            _requestLock.Release();
        }
    }

    private static HapHtmlDocument LoadDocument(string html)
    {
        var doc = new HapHtmlDocument();
        doc.LoadHtml(html);
        return doc;
    }

    private static HtmlNode? FindPostNode(HtmlNode link)
    {
        return link.Ancestors()
            .Where(node => node.Id.StartsWith("post_", StringComparison.OrdinalIgnoreCase))
            .LastOrDefault();
    }

    private static string InferAttachmentName(HtmlNode link)
    {
        var text = CleanText(link.InnerText);
        if (ArchiveExtractor.IsSupportedAttachmentName(text))
        {
            return text;
        }

        var parentText = CleanText(link.ParentNode?.InnerText ?? "");
        var match = Regex.Match(parentText, @"[^\s\\/:""<>|?*]+?\.(?:zip|7z|rar|ass|ssa|srt|sub|idx|sup|vtt|smi|ttml|dfxp)", RegexOptions.IgnoreCase);
        return match.Success ? match.Value : text;
    }

    private static string ExtractAuthor(HtmlNode? post)
    {
        if (post is null)
        {
            return "";
        }

        var author = post.SelectSingleNode(".//a[contains(concat(' ', normalize-space(@class), ' '), ' xw1 ')]")
            ?? post.SelectSingleNode(".//*[contains(@id,'favatar')]//a[normalize-space()]");
        return CleanText(author?.InnerText ?? "");
    }

    private static string ExtractPostedAt(HtmlNode? post)
    {
        if (post is null)
        {
            return "";
        }

        var node = post.SelectSingleNode(".//em[contains(@id,'authorposton')]")
            ?? post.SelectSingleNode(".//div[contains(concat(' ', normalize-space(@class), ' '), ' authi ')]//em");
        return CleanText(node?.InnerText ?? "");
    }

    private static string ExtractExcerpt(HtmlNode? post)
    {
        if (post is null)
        {
            return "";
        }

        var body = post.SelectSingleNode(".//*[starts-with(@id,'postmessage_')]") ?? post;
        var text = CleanText(body.InnerText);
        return text.Length <= 180 ? text : text[..180] + "...";
    }

    private static int ScoreSearchResult(string title, string snippet, string forumName, string query)
    {
        var score = 0;
        var combined = $"{title} {snippet}";
        var queryTokens = Tokenize(query);
        var resultTokens = Tokenize(combined);
        score += queryTokens.Intersect(resultTokens, StringComparer.OrdinalIgnoreCase).Count() * 10;
        if (combined.Contains(query, StringComparison.OrdinalIgnoreCase))
        {
            score += 18;
        }

        if (combined.Contains("字幕", StringComparison.OrdinalIgnoreCase))
        {
            score += 18;
        }

        if (PreferredForums.Contains(forumName))
        {
            score += 24;
        }

        if (DeprioritizedForums.Contains(forumName))
        {
            score -= 20;
        }

        if (Regex.IsMatch(combined, "(简|繁|chs|cht)", RegexOptions.IgnoreCase))
        {
            score += 4;
        }

        return score;
    }

    private static HashSet<string> Tokenize(string text)
    {
        return Regex.Matches(text.ToLowerInvariant(), @"[a-z0-9]+|[\u3040-\u30ff]+|[\u4e00-\u9fff]+")
            .Select(match => match.Value)
            .Where(token => token.Length > 1)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static int ParseTid(string url)
    {
        var match = Regex.Match(url, @"(?:tid=|thread-)(\d+)", RegexOptions.IgnoreCase);
        return match.Success ? int.Parse(match.Groups[1].Value) : 0;
    }

    private static string ToAbsoluteUrl(string href)
    {
        href = HapHtmlEntity.DeEntitize(href);
        return new Uri(BaseUri, href).AbsoluteUri;
    }

    private static string BuildAttachmentKey(string url)
    {
        var match = Regex.Match(url, @"[?&]aid=([^&]+)", RegexOptions.IgnoreCase);
        var source = match.Success ? WebUtility.UrlDecode(match.Groups[1].Value) : url;
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(source))).ToLowerInvariant();
    }

    private static string? ParseDispositionFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var starMatch = Regex.Match(value, @"filename\*\s*=\s*utf-8''([^;]+)", RegexOptions.IgnoreCase);
        if (starMatch.Success)
        {
            return WebUtility.UrlDecode(starMatch.Groups[1].Value.Trim('"', '\''));
        }

        var quotedMatch = Regex.Match(value, "filename\\s*=\\s*\"([^\"]+)\"", RegexOptions.IgnoreCase);
        if (quotedMatch.Success)
        {
            return WebUtility.UrlDecode(quotedMatch.Groups[1].Value);
        }

        var plainMatch = Regex.Match(value, @"filename\s*=\s*([^;]+)", RegexOptions.IgnoreCase);
        return plainMatch.Success ? WebUtility.UrlDecode(plainMatch.Groups[1].Value.Trim('"', '\'', ' ')) : null;
    }

    private static string CleanText(string text)
    {
        return Regex.Replace(HapHtmlEntity.DeEntitize(text), @"\s+", " ").Trim();
    }

    private static string SafeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(fileName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return cleaned.Length > 0 ? cleaned : "attachment.bin";
    }
}
