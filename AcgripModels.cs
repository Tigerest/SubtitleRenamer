namespace SubtitleRenamer;

internal sealed record AcgripThreadResult(
    int Tid,
    string Title,
    string Url,
    string ForumName,
    string Snippet,
    int Score);

internal sealed record AcgripAttachment(
    string Name,
    string Url,
    string Extension,
    string CacheKey);

internal sealed record AcgripThreadFloor(
    AcgripThreadResult Thread,
    int Page,
    int FloorNumber,
    string Author,
    string PostedAt,
    string Excerpt,
    IReadOnlyList<AcgripAttachment> Attachments)
{
    public string FloorLabel => $"第 {FloorNumber} 楼";
}

internal sealed record AcgripSearchSnapshot(
    string Query,
    DateTime CreatedAt,
    IReadOnlyList<AcgripThreadFloor> Floors);

internal sealed record AcgripDownloadProgress(
    string FileName,
    long BytesReceived,
    long? TotalBytes,
    double BytesPerSecond);

internal sealed class OnlineSubtitleItem
{
    public OnlineSubtitleItem(string fullPath, string threadTitle, string floorLabel, string attachmentName, string folderLabel)
    {
        FullPath = fullPath;
        ThreadTitle = threadTitle;
        FloorLabel = floorLabel;
        AttachmentName = attachmentName;
        FolderLabel = folderLabel;
    }

    public string FullPath { get; }
    public string ThreadTitle { get; }
    public string FloorLabel { get; }
    public string AttachmentName { get; }
    public string FolderLabel { get; }
    public bool Imported { get; set; }

    public string Name => Path.GetFileName(FullPath);
    public string DirectoryName => Path.GetDirectoryName(FullPath) ?? "";
}
