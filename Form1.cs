using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace SubtitleRenamer;

public partial class Form1 : Form
{
    private readonly List<FileItem> _videos = new();
    private readonly List<FileItem> _subtitles = new();
    private List<LanguageRule> _languageRules = new();
    private List<CandidatePair>? _previewSlots;
    private string _previewSignature = "";
    private int _previewDragRowIndex = -1;
    private Point _previewMouseDownPoint;
    private const string PreviewRowDataFormat = "SubtitleRenamerPreviewRow";

    private DataGridView _videoGrid = null!;
    private DataGridView _subtitleGrid = null!;
    private DataGridView _previewGrid = null!;
    private ComboBox _suffixMode = null!;
    private TextBox _suffixText = null!;
    private RadioButton _renameOriginal = null!;
    private RadioButton _createHardlink = null!;
    private CheckBox _overwriteExisting = null!;
    private CheckBox _topMost = null!;
    private Label _statusLabel = null!;
    private ToolTip _toolTip = null!;
    private OnlineSubtitleCache? _onlineSubtitleCache;
    private OnlineSubtitleForm? _onlineSubtitleForm;
    private bool _hardlinkBlockedByOnlineSubtitles;

    private static readonly HashSet<string> VideoExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".mkv", ".mp4", ".m4v", ".avi", ".mov", ".wmv", ".flv", ".webm",
        ".ts", ".m2ts", ".mts", ".mpg", ".mpeg", ".rm", ".rmvb", ".ogm", ".vob", ".3gp"
    };

    private static readonly HashSet<string> SubtitleExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".ass", ".ssa", ".srt", ".sub", ".idx", ".sup", ".vtt", ".smi", ".ttml", ".dfxp"
    };

    private static readonly HashSet<string> LanguageTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "chs", "cht", "ch", "sc", "tc", "gb", "gbk", "big5", "cn", "zh",
        "zh-cn", "zh-hans", "zh-sg", "zh-tw", "zh-hant", "zh-hk",
        "简", "繁", "简体", "繁体", "简中", "繁中", "中文", "字幕",
        "jpn", "jp", "ja", "japanese", "日", "日文", "日语",
        "eng", "en", "english", "英语", "英文",
        "chs&jpn", "cht&jpn", "sc&jp", "tc&jp", "chs-jpn", "cht-jpn",
        "chs+jpn", "cht+jpn", "简日", "繁日", "中日", "日中"
    };

    private static readonly HashSet<string> KnownGroupTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "caso", "kamigami", "dmhy", "dmg", "vcb-studio", "vcb", "qsw",
        "airota", "sumisora", "sumi-sora", "nekomoe", "lilith", "mce",
        "skymoon", "xy", "mabors", "sakurato", "lolihouse", "haruhana"
    };

    private static readonly HashSet<string> IgnoredSuffixTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "1080p", "720p", "2160p", "480p", "4k", "8k", "bdrip", "bdrips",
        "bluray", "blu-ray", "web-dl", "webrip", "web", "tv", "dvdrip",
        "hevc", "avc", "h264", "h265", "x264", "x265", "10bit", "8bit",
        "hi10p", "ma10p", "aac", "flac", "dts", "opus", "mp3", "hdr", "sdr",
        "ncop", "nced", "ova", "oad", "sp"
    };

    [DllImport("Kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

    public Form1()
    {
        InitializeComponent();
        _languageRules = LoadLanguageRules();
        RegisterLanguageTokens(_languageRules);
        BuildUi();
        FormClosed += (_, _) => CleanupOnlineSubtitleCache();
        RefreshAll("拖入视频和字幕文件，或使用上方按钮导入。");
    }

    private void BuildUi()
    {
        Text = "大河字幕重命名工具";
        var appIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        if (appIcon is not null)
        {
            Icon = appIcon;
        }

        MinimumSize = new Size(1220, 760);
        Size = new Size(1800, 900);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(246, 247, 249);
        AllowDrop = true;
        DragEnter += HandleDragEnter;
        DragDrop += HandleDragDrop;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(12),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 48));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 52));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 96));
        Controls.Add(root);
        _toolTip = new ToolTip { ShowAlways = true };

        root.Controls.Add(BuildToolbar(), 0, 0);
        root.Controls.Add(BuildListsArea(), 0, 1);
        root.Controls.Add(BuildPreviewArea(), 0, 2);
        root.Controls.Add(BuildOptionsArea(), 0, 3);
    }

    private Control BuildToolbar()
    {
        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 3, 0, 3),
            BackColor = BackColor
        };

        panel.Controls.Add(CreateButton("导入视频", (_, _) => ImportByDialog(true)));
        panel.Controls.Add(CreateButton("导入字幕", (_, _) => ImportByDialog(false)));
        panel.Controls.Add(CreateButton("导入文件夹", (_, _) => ImportFolder()));
        panel.Controls.Add(CreateButton("在线查找", (_, _) => OpenOnlineSubtitleSearch()));
        panel.Controls.Add(CreateButton("语言后缀", (_, _) => OpenLanguageSettings()));
        panel.Controls.Add(CreateButton("清空全部", (_, _) => ClearAll()));
        panel.Controls.Add(CreateButton("预览重置", (_, _) => ResetPreview("预览调整已重置。")));

        var applyButton = CreateButton("应用更名", (_, _) => ApplyRenames());
        applyButton.BackColor = Color.FromArgb(30, 120, 212);
        applyButton.ForeColor = Color.White;
        panel.Controls.Add(applyButton);

        panel.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "  可直接把视频、字幕或文件夹拖到窗口里，支持一次匹配多语言字幕",
            ForeColor = Color.FromArgb(86, 96, 112),
            Padding = new Padding(8, 10, 0, 0)
        });

        return panel;
    }

    private Control BuildListsArea()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            BackColor = BackColor
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        _videoGrid = CreateFileGrid();
        _subtitleGrid = CreateFileGrid();
        table.Controls.Add(BuildFileGroup("视频文件列表", _videoGrid, true), 0, 0);
        table.Controls.Add(BuildFileGroup("字幕文件列表", _subtitleGrid, false), 1, 0);
        return table;
    }

    private Control BuildFileGroup(string title, DataGridView grid, bool videoList)
    {
        var group = new GroupBox
        {
            Text = title,
            Dock = DockStyle.Fill,
            Padding = new Padding(10),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(34, 40, 49)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.White
        };
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
        layout.Controls.Add(grid, 0, 0);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 6, 0, 4),
            BackColor = Color.White
        };
        if (videoList)
        {
            buttons.Controls.Add(CreateSmallButton("上移", (_, _) => MoveSelected(videoList, -1)));
            buttons.Controls.Add(CreateSmallButton("下移", (_, _) => MoveSelected(videoList, 1)));
        }

        buttons.Controls.Add(CreateSmallButton("删除", (_, _) => RemoveSelected(videoList)));
        if (videoList)
        {
            buttons.Controls.Add(CreateSmallButton("重置排序", (_, _) => SortList(videoList)));
        }

        buttons.Controls.Add(CreateSmallButton(videoList ? "清空视频" : "清空字幕", (_, _) => ClearList(videoList)));
        layout.Controls.Add(buttons, 0, 1);

        group.Controls.Add(layout);
        return group;
    }

    private Control BuildPreviewArea()
    {
        var group = new GroupBox
        {
            Text = "预览匹配与目标文件名",
            Dock = DockStyle.Fill,
            Padding = new Padding(10),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(34, 40, 49)
        };

        _previewGrid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
            EnableHeadersVisualStyles = false,
            GridColor = Color.FromArgb(231, 235, 241),
            MultiSelect = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };
        StyleGrid(_previewGrid);
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "#", FillWeight = 8 });
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "视频", FillWeight = 34 });
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "字幕", FillWeight = 34 });
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "识别/使用后缀", FillWeight = 18 });
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "重命名为", FillWeight = 46 });
        _previewGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", FillWeight = 18 });
        _previewGrid.AllowDrop = true;
        _previewGrid.MouseDown += HandlePreviewMouseDown;
        _previewGrid.MouseMove += HandlePreviewMouseMove;
        _previewGrid.DragEnter += HandlePreviewDragEnter;
        _previewGrid.DragOver += HandlePreviewDragOver;
        _previewGrid.DragDrop += HandlePreviewDragDrop;

        group.Controls.Add(_previewGrid);
        return group;
    }

    private Control BuildOptionsArea()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = BackColor,
            Padding = new Padding(0, 10, 0, 0)
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));

        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            BackColor = BackColor,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0)
        };

        panel.Controls.Add(CreateOptionLabel("后缀策略"));

        _suffixMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 230,
            Margin = new Padding(0, 0, 18, 0)
        };
        _suffixMode.Items.AddRange(new object[]
        {
            "识别原语言后缀",
            "统一使用默认后缀",
            "默认后缀+语言后缀",
            "不添加后缀"
        });
        _suffixMode.SelectedIndex = 0;
        _suffixMode.SelectedIndexChanged += (_, _) => RefreshAll("后缀策略已更新。");
        panel.Controls.Add(_suffixMode);

        panel.Controls.Add(CreateOptionLabel("默认后缀"));

        _suffixText = new TextBox { Text = ".chs", Width = 118, Margin = new Padding(0, 0, 18, 0) };
        _suffixText.TextChanged += (_, _) => RefreshAll("默认后缀已更新。");
        panel.Controls.Add(_suffixText);

        _renameOriginal = new RadioButton
        {
            Text = "重命名原字幕",
            Checked = true,
            AutoSize = true,
            Margin = new Padding(0, 4, 18, 0)
        };
        panel.Controls.Add(_renameOriginal);

        _createHardlink = new RadioButton
        {
            Text = "硬链接副本，保留原字幕",
            AutoSize = true,
            Margin = new Padding(0, 4, 18, 0)
        };
        _createHardlink.Click += (_, _) =>
        {
            if (!_hardlinkBlockedByOnlineSubtitles)
            {
                return;
            }

            _renameOriginal.Checked = true;
            _statusLabel.Text = "在线搜索模式下不支持硬链接。";
        };
        panel.Controls.Add(_createHardlink);

        var rightOptions = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Margin = new Padding(0)
        };
        _overwriteExisting = new CheckBox
        {
            Text = "覆盖已有同名字幕",
            AutoSize = true,
            Margin = new Padding(0, 4, 18, 0)
        };
        _overwriteExisting.CheckedChanged += (_, _) => RefreshAll("覆盖策略已更新。");
        rightOptions.Controls.Add(_overwriteExisting);

        _statusLabel = new Label
        {
            AutoSize = true,
            ForeColor = Color.FromArgb(86, 96, 112),
            Margin = new Padding(0, 5, 18, 0)
        };
        rightOptions.Controls.Add(_statusLabel);
        panel.Controls.Add(rightOptions);

        _topMost = new CheckBox
        {
            Text = "置顶",
            AutoSize = true,
            Margin = new Padding(0, 4, 0, 0)
        };
        _topMost.CheckedChanged += (_, _) => TopMost = _topMost.Checked;
        panel.Controls.Add(_topMost);

        var bilibiliLink = new LinkLabel
        {
            Text = "https://space.bilibili.com/12562485",
            AutoSize = true,
            Font = new Font(Font.FontFamily, 8F, FontStyle.Regular, GraphicsUnit.Point),
            LinkColor = Color.FromArgb(30, 120, 212),
            ActiveLinkColor = Color.FromArgb(20, 88, 156),
            VisitedLinkColor = Color.FromArgb(30, 120, 212),
            Margin = new Padding(0, 2, 0, 0),
            LinkBehavior = LinkBehavior.HoverUnderline
        };
        bilibiliLink.LinkClicked += (_, _) => OpenBilibiliHome();

        root.Controls.Add(panel, 0, 0);
        root.Controls.Add(bilibiliLink, 0, 1);
        return root;
    }

    private static void OpenBilibiliHome()
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = "https://space.bilibili.com/12562485",
            UseShellExecute = true
        });
    }

    private static Label CreateOptionLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            Margin = new Padding(0, 5, 8, 0),
            UseMnemonic = false
        };
    }

    private static Button CreateButton(string text, EventHandler handler)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            MinimumSize = new Size(96, 34),
            Margin = new Padding(0, 0, 8, 0),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(32, 38, 46),
            UseVisualStyleBackColor = false
        };
        button.FlatAppearance.BorderColor = Color.FromArgb(204, 211, 222);
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(238, 244, 252);
        button.Click += handler;
        return button;
    }

    private static Button CreateSmallButton(string text, EventHandler handler)
    {
        var button = CreateButton(text, handler);
        button.AutoSize = false;
        button.Size = text.Length > 3 ? new Size(132, 34) : new Size(88, 34);
        button.MinimumSize = button.Size;
        button.Margin = new Padding(0, 0, 6, 0);
        return button;
    }

    private DataGridView CreateFileGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
            EnableHeadersVisualStyles = false,
            GridColor = Color.FromArgb(231, 235, 241),
            MultiSelect = false,
            ReadOnly = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };
        StyleGrid(grid);
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "#", FillWeight = 10 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "文件名", FillWeight = 56 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "目录", FillWeight = 44 });
        grid.AllowDrop = true;
        grid.DragEnter += HandleDragEnter;
        grid.DragDrop += HandleDragDrop;
        return grid;
    }

    private static void StyleGrid(DataGridView grid)
    {
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(242, 245, 249);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(45, 52, 64);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
        grid.ColumnHeadersHeight = 32;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(218, 234, 252);
        grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(20, 28, 38);
        grid.DefaultCellStyle.ForeColor = Color.FromArgb(34, 40, 49);
        grid.RowTemplate.Height = 28;
    }

    private void ImportByDialog(bool videos)
    {
        using var dialog = new OpenFileDialog
        {
            Title = videos ? "选择视频文件" : "选择字幕文件",
            Multiselect = true,
            Filter = videos
                ? "视频文件|*.mkv;*.mp4;*.m4v;*.avi;*.mov;*.wmv;*.flv;*.webm;*.ts;*.m2ts;*.mts;*.mpg;*.mpeg;*.rm;*.rmvb;*.ogm;*.vob;*.3gp|所有文件|*.*"
                : "字幕文件|*.ass;*.ssa;*.srt;*.sub;*.idx;*.sup;*.vtt;*.smi;*.ttml;*.dfxp|所有文件|*.*"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var stats = AddPaths(dialog.FileNames, videos ? ImportKind.VideoOnly : ImportKind.SubtitleOnly);
        RefreshAll(BuildImportMessage(stats));
    }

    private void ImportFolder()
    {
        using var dialog = new FolderBrowserDialog
        {
            Description = "选择包含视频和字幕的文件夹",
            UseDescriptionForTitle = true
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var stats = AddPaths(new[] { dialog.SelectedPath }, ImportKind.Mixed);
        RefreshAll(BuildImportMessage(stats));
    }

    private void OpenOnlineSubtitleSearch()
    {
        _onlineSubtitleCache ??= new OnlineSubtitleCache();
        if (_onlineSubtitleForm is null || _onlineSubtitleForm.IsDisposed)
        {
            _onlineSubtitleForm = new OnlineSubtitleForm(
                _onlineSubtitleCache,
                GuessDefaultOnlineSearchQuery,
                ImportOnlineSubtitles);
        }

        _onlineSubtitleForm.PrepareDefaultQuery();
        if (!_onlineSubtitleForm.Visible)
        {
            _onlineSubtitleForm.Show(this);
        }

        _onlineSubtitleForm.Activate();
    }

    private string GuessDefaultOnlineSearchQuery()
    {
        return _videos.Count == 0 ? "" : TitleGuess.GuessFromVideoName(_videos[0].Name);
    }

    private void ImportOnlineSubtitles(IReadOnlyList<string> paths)
    {
        var stats = AddPaths(paths, ImportKind.SubtitleOnly, FileSource.Online);
        ResetPreview($"在线搜索导入：字幕 {stats.Subtitles} 个，重复 {stats.Duplicates} 个，排除 {stats.Ignored} 个。");
    }

    private void CleanupOnlineSubtitleCache()
    {
        try
        {
            if (_onlineSubtitleForm is not null && !_onlineSubtitleForm.IsDisposed)
            {
                _onlineSubtitleForm.CloseForApplicationExit();
                _onlineSubtitleForm.Dispose();
            }
        }
        catch
        {
            // Closing the main window should not be blocked by an auxiliary window.
        }

        _onlineSubtitleCache?.Dispose();
        _onlineSubtitleCache = null;
    }

    private void HandleDragEnter(object? sender, DragEventArgs e)
    {
        e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true
            ? DragDropEffects.Copy
            : DragDropEffects.None;
    }

    private void HandleDragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is not string[] paths)
        {
            return;
        }

        var stats = AddPaths(paths, ImportKind.Mixed);
        RefreshAll(BuildImportMessage(stats));
    }

    private void HandlePreviewMouseDown(object? sender, MouseEventArgs e)
    {
        _previewDragRowIndex = -1;
        if (e.Button != MouseButtons.Left)
        {
            return;
        }

        var hit = _previewGrid.HitTest(e.X, e.Y);
        if (hit.RowIndex < 0)
        {
            return;
        }

        EnsurePreviewSlots();
        if (_previewSlots is null || hit.RowIndex >= _previewSlots.Count)
        {
            return;
        }

        if (_previewSlots[hit.RowIndex].Subtitle is not null)
        {
            return;
        }

        _previewDragRowIndex = hit.RowIndex;
        _previewMouseDownPoint = e.Location;
    }

    private void HandlePreviewMouseMove(object? sender, MouseEventArgs e)
    {
        if ((e.Button & MouseButtons.Left) == 0 || _previewDragRowIndex < 0)
        {
            return;
        }

        var dragSize = SystemInformation.DragSize;
        var dragBounds = new Rectangle(
            _previewMouseDownPoint.X - dragSize.Width / 2,
            _previewMouseDownPoint.Y - dragSize.Height / 2,
            dragSize.Width,
            dragSize.Height);

        if (dragBounds.Contains(e.Location))
        {
            return;
        }

        var data = new DataObject();
        data.SetData(PreviewRowDataFormat, _previewDragRowIndex);
        _previewGrid.DoDragDrop(data, DragDropEffects.Move);
    }

    private void HandlePreviewDragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(PreviewRowDataFormat) == true)
        {
            e.Effect = DragDropEffects.Move;
            return;
        }

        HandleDragEnter(sender, e);
    }

    private void HandlePreviewDragOver(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(PreviewRowDataFormat) != true)
        {
            HandleDragEnter(sender, e);
            return;
        }

        e.Effect = TryGetPreviewDropRows(e, out var sourceIndex, out var targetIndex) &&
                   CanDropPreviewRow(sourceIndex, targetIndex)
            ? DragDropEffects.Move
            : DragDropEffects.None;
    }

    private void HandlePreviewDragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(PreviewRowDataFormat) != true)
        {
            HandleDragDrop(sender, e);
            return;
        }

        if (!TryGetPreviewDropRows(e, out var sourceIndex, out var targetIndex) ||
            !CanDropPreviewRow(sourceIndex, targetIndex))
        {
            _statusLabel.Text = "只能拖动空行调整字幕缺失位置。";
            return;
        }

        ApplyPreviewDrop(sourceIndex, targetIndex);
        RefreshAll("预览匹配已调整。");
    }

    private void EnsurePreviewSlots()
    {
        var signature = BuildPreviewSignature();
        if (_previewSlots is null || !string.Equals(_previewSignature, signature, StringComparison.Ordinal))
        {
            _previewSlots = BuildCandidatePairs();
            _previewSignature = signature;
        }
    }

    private bool TryGetPreviewDropRows(DragEventArgs e, out int sourceIndex, out int targetIndex)
    {
        sourceIndex = -1;
        targetIndex = -1;
        if (e.Data?.GetData(PreviewRowDataFormat) is not int rowIndex)
        {
            return false;
        }

        var point = _previewGrid.PointToClient(new Point(e.X, e.Y));
        var hit = _previewGrid.HitTest(point.X, point.Y);
        if (hit.RowIndex < 0)
        {
            return false;
        }

        EnsurePreviewSlots();
        if (_previewSlots is null || rowIndex < 0 || rowIndex >= _previewSlots.Count || hit.RowIndex >= _previewSlots.Count)
        {
            return false;
        }

        sourceIndex = rowIndex;
        targetIndex = hit.RowIndex;
        return sourceIndex != targetIndex;
    }

    private bool CanDropPreviewRow(int sourceIndex, int targetIndex)
    {
        if (_previewSlots is null)
        {
            return false;
        }

        return sourceIndex != targetIndex && _previewSlots[sourceIndex].Subtitle is null;
    }

    private void ApplyPreviewDrop(int sourceIndex, int targetIndex)
    {
        if (_previewSlots is null)
        {
            return;
        }

        var subtitles = _previewSlots.Select(slot => slot.Subtitle).ToList();
        subtitles.RemoveAt(sourceIndex);
        subtitles.Insert(targetIndex, null);

        for (var i = 0; i < _previewSlots.Count; i++)
        {
            var subtitle = subtitles[i];
            _previewSlots[i] = _previewSlots[i] with
            {
                Subtitle = subtitle,
                SlotSuffix = subtitle is null ? _previewSlots[i].SlotSuffix : GetEffectiveSuffixForGrouping(subtitle)
            };
        }
    }

    private ImportStats AddPaths(IEnumerable<string> paths, ImportKind kind, FileSource source = FileSource.Local)
    {
        var stats = new ImportStats();
        foreach (var path in ExpandPaths(paths))
        {
            var extension = Path.GetExtension(path);
            if (VideoExtensions.Contains(extension) && kind != ImportKind.SubtitleOnly)
            {
                if (AddUnique(_videos, path, FileSource.Local))
                {
                    stats.Videos++;
                }
                else
                {
                    stats.Duplicates++;
                }
            }
            else if (SubtitleExtensions.Contains(extension) && kind != ImportKind.VideoOnly)
            {
                if (AddUnique(_subtitles, path, source))
                {
                    stats.Subtitles++;
                }
                else
                {
                    stats.Duplicates++;
                }
            }
            else
            {
                stats.Ignored++;
            }
        }

        _videos.Sort(FileItem.CompareByName);
        _subtitles.Sort(CompareSubtitleByCanonicalName);
        return stats;
    }

    private static IEnumerable<string> ExpandPaths(IEnumerable<string> paths)
    {
        foreach (var path in paths)
        {
            if (File.Exists(path))
            {
                yield return Path.GetFullPath(path);
                continue;
            }

            if (!Directory.Exists(path))
            {
                continue;
            }

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories);
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                yield return Path.GetFullPath(file);
            }
        }
    }

    private static bool AddUnique(List<FileItem> list, string path, FileSource source = FileSource.Local)
    {
        var fullPath = Path.GetFullPath(path);
        if (list.Any(item => string.Equals(item.FullPath, fullPath, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        list.Add(new FileItem(fullPath, source));
        return true;
    }

    private void ClearAll()
    {
        if (_videos.Count + _subtitles.Count > 0)
        {
            var result = MessageBox.Show(this, "清空当前导入的视频和字幕列表？", "确认清空",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (result != DialogResult.OK)
            {
                return;
            }
        }

        _videos.Clear();
        _subtitles.Clear();
        RefreshAll("列表已清空。");
    }

    private void OpenLanguageSettings()
    {
        using var dialog = new LanguageSettingsDialog(CloneRules(_languageRules));
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        _languageRules = SanitizeRules(dialog.Rules);
        if (_languageRules.Count == 0)
        {
            _languageRules = DefaultLanguageRules();
        }

        RegisterLanguageTokens(_languageRules);
        SaveLanguageRules(_languageRules);
        ResetPreview("语言后缀设置已保存，预览已重置。");
    }

    private void ResetLanguageSettings()
    {
        var result = MessageBox.Show(this, "恢复默认语言后缀设置？", "确认恢复默认",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
        if (result != DialogResult.OK)
        {
            return;
        }

        _languageRules = DefaultLanguageRules();
        RegisterLanguageTokens(_languageRules);
        SaveLanguageRules(_languageRules);
        ResetPreview("语言后缀设置已恢复默认，预览已重置。");
    }

    private void SortList(bool videoList)
    {
        var list = videoList ? _videos : _subtitles;
        if (videoList)
        {
            list.Sort(FileItem.CompareByName);
        }
        else
        {
            list.Sort(CompareSubtitleByCanonicalName);
        }

        ResetPreview(videoList ? "视频列表已重置排序，预览已重置。" : "字幕列表已重置排序，预览已重置。");
    }

    private int CompareSubtitleByCanonicalName(FileItem left, FileItem right)
    {
        var leftKey = BuildCanonicalSubtitleSortKey(left.FullPath);
        var rightKey = BuildCanonicalSubtitleSortKey(right.FullPath);
        var nameCompare = StringComparer.CurrentCultureIgnoreCase.Compare(leftKey, rightKey);
        return nameCompare != 0
            ? nameCompare
            : FileItem.CompareByName(left, right);
    }

    private string BuildCanonicalSubtitleSortKey(string path)
    {
        var stem = Path.GetFileNameWithoutExtension(path);
        var extension = Path.GetExtension(path);
        var detected = DetectSuffixFromSubtitleName(path);
        var canonicalLanguageSuffix = CanonicalizeLanguageSuffix(detected);
        if (detected.Length == 0 || canonicalLanguageSuffix.Length == 0)
        {
            return Path.GetFileName(path);
        }

        var baseStem = TrimDetectedTailSuffix(stem, detected);
        return baseStem + canonicalLanguageSuffix + extension;
    }

    private static string TrimDetectedTailSuffix(string stem, string detectedSuffix)
    {
        var text = stem;
        var tokens = detectedSuffix
            .TrimStart('.')
            .Split('.', StringSplitOptions.RemoveEmptyEntries)
            .Reverse();

        foreach (var token in tokens)
        {
            text = TrimOneTailToken(text, token);
        }

        return text.TrimEnd('.', ' ', '_', '-');
    }

    private static string TrimOneTailToken(string text, string token)
    {
        var escaped = Regex.Escape(token);
        var bracketMatch = Regex.Match(text, $@"(?:[\[\(【（]\s*{escaped}\s*[\]\)】）])\s*$", RegexOptions.IgnoreCase);
        if (bracketMatch.Success)
        {
            return text[..bracketMatch.Index].TrimEnd('.', ' ', '_', '-');
        }

        var delimiterMatch = Regex.Match(text, $@"(?<prefix>.*?)[\._\-\s]+{escaped}\s*$", RegexOptions.IgnoreCase);
        if (delimiterMatch.Success)
        {
            return delimiterMatch.Groups["prefix"].Value.TrimEnd('.', ' ', '_', '-');
        }

        if (text.Length == token.Length && string.Equals(text, token, StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        return text;
    }

    private void ClearList(bool videoList)
    {
        var list = videoList ? _videos : _subtitles;
        if (list.Count == 0)
        {
            return;
        }

        var label = videoList ? "视频" : "字幕";
        var result = MessageBox.Show(this, $"清空当前{label}文件列表？", $"确认清空{label}列表",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
        if (result != DialogResult.OK)
        {
            return;
        }

        list.Clear();
        RefreshAll($"{label}列表已清空。");
    }

    private void MoveSelected(bool videoList, int direction)
    {
        var grid = videoList ? _videoGrid : _subtitleGrid;
        var list = videoList ? _videos : _subtitles;
        if (grid.CurrentRow is null || grid.CurrentRow.Index < 0)
        {
            return;
        }

        var oldIndex = grid.CurrentRow.Index;
        var newIndex = oldIndex + direction;
        if (newIndex < 0 || newIndex >= list.Count)
        {
            return;
        }

        (list[oldIndex], list[newIndex]) = (list[newIndex], list[oldIndex]);
        RefreshAll(videoList ? "视频顺序已调整。" : "字幕顺序已调整。");
        SelectRow(grid, newIndex);
    }

    private void RemoveSelected(bool videoList)
    {
        var grid = videoList ? _videoGrid : _subtitleGrid;
        var list = videoList ? _videos : _subtitles;
        if (grid.CurrentRow is null || grid.CurrentRow.Index < 0)
        {
            return;
        }

        var index = grid.CurrentRow.Index;
        list.RemoveAt(index);
        RefreshAll(videoList ? "已从视频列表移除。" : "已从字幕列表移除。");
        if (list.Count > 0)
        {
            SelectRow(grid, Math.Min(index, list.Count - 1));
        }
    }

    private static void SelectRow(DataGridView grid, int index)
    {
        if (index < 0 || index >= grid.Rows.Count)
        {
            return;
        }

        grid.ClearSelection();
        grid.Rows[index].Selected = true;
        grid.CurrentCell = grid.Rows[index].Cells[0];
    }

    private void RefreshAll(string? message = null)
    {
        FillFileGrid(_videoGrid, _videos);
        FillFileGrid(_subtitleGrid, _subtitles);
        FillPreviewGrid();
        UpdateHardlinkAvailability();
        if (!string.IsNullOrWhiteSpace(message))
        {
            _statusLabel.Text = message;
        }
    }

    private static void FillFileGrid(DataGridView grid, IReadOnlyList<FileItem> items)
    {
        grid.Rows.Clear();
        for (var i = 0; i < items.Count; i++)
        {
            var name = items[i].Source == FileSource.Online ? "[在线] " + items[i].Name : items[i].Name;
            grid.Rows.Add(i + 1, name, items[i].DirectoryName);
        }
    }

    private void UpdateHardlinkAvailability()
    {
        if (_createHardlink is null || _renameOriginal is null)
        {
            return;
        }

        var hasOnlineSubtitle = _subtitles.Any(item => item.Source == FileSource.Online);
        _hardlinkBlockedByOnlineSubtitles = hasOnlineSubtitle;
        _createHardlink.AutoCheck = !hasOnlineSubtitle;
        _createHardlink.ForeColor = hasOnlineSubtitle ? Color.FromArgb(150, 150, 150) : Color.FromArgb(32, 38, 46);
        _toolTip.SetToolTip(_createHardlink, hasOnlineSubtitle ? "在线搜索模式下不支持硬链接" : "");
        if (hasOnlineSubtitle && _createHardlink.Checked)
        {
            _renameOriginal.Checked = true;
        }
    }

    private void FillPreviewGrid()
    {
        _previewGrid.Rows.Clear();
        var rows = BuildPreviewRows();

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];

            var rowIndex = _previewGrid.Rows.Add(
                i + 1,
                row.Video?.Name ?? "",
                row.Subtitle?.Name ?? (row.Video is not null ? "拖动此行调整字幕缺失位置" : ""),
                row.Suffix,
                row.TargetPath.Length > 0 ? Path.GetFileName(row.TargetPath) : "",
                row.State);

            if (row.State != "就绪" && row.State != "已同名")
            {
                _previewGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.FromArgb(176, 80, 0);
            }
        }
    }

    private List<PreviewRow> BuildPreviewRows()
    {
        var signature = BuildPreviewSignature();
        if (_previewSlots is null || !string.Equals(_previewSignature, signature, StringComparison.Ordinal))
        {
            _previewSlots = BuildCandidatePairs();
            _previewSignature = signature;
        }

        var pairs = _previewSlots;
        var targetCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var pair in pairs)
        {
            if (pair.Video is null || pair.Subtitle is null)
            {
                continue;
            }

            var suffix = GetEffectiveSuffix(pair.Video, pair.Subtitle);
            var target = BuildTargetPath(pair.Video, pair.Subtitle, suffix);
            targetCounts[target] = targetCounts.GetValueOrDefault(target) + 1;
        }

        return pairs.Select(pair =>
        {
            var suffix = pair.Subtitle is not null && pair.Video is not null
                ? GetEffectiveSuffix(pair.Video, pair.Subtitle)
                : pair.SlotSuffix;
            var target = pair.Video is not null && pair.Subtitle is not null
                ? BuildTargetPath(pair.Video, pair.Subtitle, suffix)
                : "";
            var duplicated = target.Length > 0 && targetCounts.GetValueOrDefault(target) > 1;
            var state = GetPreviewState(pair.Video, pair.Subtitle, target, duplicated);
            return new PreviewRow(pair.Video, pair.Subtitle, suffix, target, state);
        }).ToList();
    }

    private string BuildPreviewSignature()
    {
        var builder = new StringBuilder()
            .Append(_suffixMode.SelectedIndex)
            .Append('|')
            .Append(NormalizeSuffix(_suffixText.Text));

        foreach (var video in _videos)
        {
            builder.Append("|v:").Append(video.FullPath);
        }

        foreach (var subtitle in _subtitles)
        {
            builder.Append("|s:").Append(subtitle.FullPath);
        }

        return builder.ToString();
    }

    private void ResetPreview(string message)
    {
        _previewSlots = null;
        _previewSignature = "";
        RefreshAll(message);
    }

    private List<CandidatePair> BuildCandidatePairs()
    {
        var pairs = new List<CandidatePair>();
        if (_videos.Count == 0 && _subtitles.Count == 0)
        {
            return pairs;
        }

        if (_videos.Count == 0)
        {
            pairs.AddRange(_subtitles.Select(subtitle => new CandidatePair(null, subtitle)));
            return pairs;
        }

        if (_subtitles.Count == 0)
        {
            pairs.AddRange(_videos.Select(video => new CandidatePair(video, null)));
            return pairs;
        }

        var sortedSubtitles = _subtitles
            .OrderBy(item => item, Comparer<FileItem>.Create(CompareSubtitleByCanonicalName))
            .ToList();
        var distinctSuffixCount = sortedSubtitles
            .Select(GetEffectiveSuffixForGrouping)
            .Where(suffix => suffix.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();
        var useMultiSubtitleMode = sortedSubtitles.Count > _videos.Count || distinctSuffixCount > 1;
        var suffixOrder = BuildSuffixOrder(sortedSubtitles);

        if (useMultiSubtitleMode)
        {
            if (sortedSubtitles.Count % _videos.Count == 0)
            {
                return BuildCompleteMultiSubtitlePairs(sortedSubtitles, suffixOrder);
            }

            if (suffixOrder.Count > 1)
            {
                return BuildOrderedSlotPairs(sortedSubtitles, suffixOrder);
            }

            return BuildSequentialPairs(sortedSubtitles);
        }

        return BuildSequentialPairs(_subtitles);
    }

    private List<CandidatePair> BuildCompleteMultiSubtitlePairs(IReadOnlyList<FileItem> sortedSubtitles, IReadOnlyList<string> suffixOrder)
    {
        var tracksPerVideo = sortedSubtitles.Count / _videos.Count;
        if (tracksPerVideo <= 1)
        {
            return BuildSequentialPairs(sortedSubtitles);
        }

        if (LooksLikeInterleavedLanguageBlocks(sortedSubtitles, tracksPerVideo))
        {
            if (suffixOrder.Count == tracksPerVideo)
            {
                return BuildOrderedSlotPairs(sortedSubtitles, suffixOrder);
            }
        }

        return BuildContiguousTrackPairs(sortedSubtitles, tracksPerVideo);
    }

    private bool LooksLikeInterleavedLanguageBlocks(IReadOnlyList<FileItem> sortedSubtitles, int blockSize)
    {
        if (blockSize <= 1 || sortedSubtitles.Count != _videos.Count * blockSize)
        {
            return false;
        }

        for (var offset = 0; offset < sortedSubtitles.Count; offset += blockSize)
        {
            var suffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < blockSize; i++)
            {
                var suffix = GetEffectiveSuffixForGrouping(sortedSubtitles[offset + i]);
                if (suffix.Length == 0 || !suffixes.Add(suffix))
                {
                    return false;
                }
            }
        }

        return true;
    }

    private List<CandidatePair> BuildContiguousTrackPairs(IReadOnlyList<FileItem> sortedSubtitles, int trackCount)
    {
        var pairs = new List<CandidatePair>();
        for (var trackIndex = 0; trackIndex < trackCount; trackIndex++)
        {
            var offset = trackIndex * _videos.Count;
            for (var videoIndex = 0; videoIndex < _videos.Count; videoIndex++)
            {
                var subtitle = sortedSubtitles[offset + videoIndex];
                pairs.Add(new CandidatePair(_videos[videoIndex], subtitle, GetEffectiveSuffixForGrouping(subtitle)));
            }
        }

        return pairs;
    }

    private List<CandidatePair> BuildOrderedSlotPairs(IReadOnlyList<FileItem> sortedSubtitles, IReadOnlyList<string> suffixOrder)
    {
        var slotCount = _videos.Count * suffixOrder.Count;
        var pairs = new List<CandidatePair>(slotCount);
        for (var slotIndex = 0; slotIndex < slotCount; slotIndex++)
        {
            pairs.Add(new CandidatePair(_videos[slotIndex / suffixOrder.Count], null, suffixOrder[slotIndex % suffixOrder.Count]));
        }

        var currentSlot = 0;
        foreach (var subtitle in sortedSubtitles)
        {
            var suffix = GetEffectiveSuffixForGrouping(subtitle);
            var suffixIndex = IndexOfSuffix(suffixOrder, suffix);

            if (suffixIndex >= 0)
            {
                while (currentSlot < pairs.Count && currentSlot % suffixOrder.Count != suffixIndex)
                {
                    currentSlot++;
                }
            }

            if (currentSlot < pairs.Count)
            {
                pairs[currentSlot] = pairs[currentSlot] with { Subtitle = subtitle };
                currentSlot++;
            }
            else
            {
                pairs.Add(new CandidatePair(null, subtitle, suffix));
            }
        }

        return pairs;
    }

    private List<string> BuildSuffixOrder(IReadOnlyList<FileItem> sortedSubtitles)
    {
        var suffixes = sortedSubtitles
            .Select(GetEffectiveSuffixForGrouping)
            .Where(suffix => suffix.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (suffixes.Count == 0)
        {
            return suffixes;
        }

        var ordered = new List<string>();
        foreach (var rule in _languageRules)
        {
            var suffix = NormalizeSuffix(rule.OutputSuffix);
            if (suffix.Length > 0 && suffixes.Any(item => string.Equals(item, suffix, StringComparison.OrdinalIgnoreCase)))
            {
                ordered.Add(suffix);
            }
        }

        foreach (var suffix in suffixes)
        {
            if (!ordered.Any(item => string.Equals(item, suffix, StringComparison.OrdinalIgnoreCase)))
            {
                ordered.Add(suffix);
            }
        }

        return ordered;
    }

    private static int IndexOfSuffix(IReadOnlyList<string> suffixes, string suffix)
    {
        for (var i = 0; i < suffixes.Count; i++)
        {
            if (string.Equals(suffixes[i], suffix, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    private BlockPairResult BuildBlockSubtitlePairs(IReadOnlyList<FileItem> sortedSubtitles)
    {
        var blocks = new List<List<FileItem>>();
        var currentBlock = new List<FileItem>();
        var currentSuffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var subtitle in sortedSubtitles)
        {
            var suffix = GetEffectiveSuffixForGrouping(subtitle);
            if (currentBlock.Count > 0 && suffix.Length > 0 && currentSuffixes.Contains(suffix))
            {
                blocks.Add(currentBlock);
                currentBlock = new List<FileItem>();
                currentSuffixes.Clear();
            }

            currentBlock.Add(subtitle);
            if (suffix.Length > 0)
            {
                currentSuffixes.Add(suffix);
            }
        }

        if (currentBlock.Count > 0)
        {
            blocks.Add(currentBlock);
        }

        var pairs = new List<CandidatePair>();
        for (var blockIndex = 0; blockIndex < blocks.Count; blockIndex++)
        {
            var video = blockIndex < _videos.Count ? _videos[blockIndex] : null;
            foreach (var subtitle in blocks[blockIndex])
            {
                pairs.Add(new CandidatePair(video, subtitle));
            }
        }

        return new BlockPairResult(blocks.Count, pairs);
    }

    private List<CandidatePair> BuildSequentialPairs(IReadOnlyList<FileItem> subtitles)
    {
        var pairs = new List<CandidatePair>();
        var count = Math.Max(_videos.Count, subtitles.Count);
        for (var i = 0; i < count; i++)
        {
            pairs.Add(new CandidatePair(
                i < _videos.Count ? _videos[i] : null,
                i < subtitles.Count ? subtitles[i] : null,
                i < subtitles.Count ? GetEffectiveSuffixForGrouping(subtitles[i]) : ""));
        }

        return pairs;
    }

    private List<SubtitleGroup> BuildSubtitleGroups()
    {
        var groups = new List<SubtitleGroup>();
        foreach (var subtitle in _subtitles)
        {
            var suffix = GetEffectiveSuffixForGrouping(subtitle);
            var group = groups.FirstOrDefault(item => string.Equals(item.Suffix, suffix, StringComparison.OrdinalIgnoreCase));
            if (group is null)
            {
                group = new SubtitleGroup(suffix);
                groups.Add(group);
            }

            group.Items.Add(subtitle);
        }

        return groups;
    }

    private string GetPreviewState(FileItem? video, FileItem? subtitle, string target, bool duplicatedTarget)
    {
        if (video is null)
        {
            return "缺少视频";
        }

        if (subtitle is null)
        {
            return "缺少字幕";
        }

        if (duplicatedTarget)
        {
            return "目标重复";
        }

        if (SamePath(subtitle.FullPath, target))
        {
            return "已同名";
        }

        if (File.Exists(target) && !_overwriteExisting.Checked)
        {
            return "目标已存在";
        }

        return "就绪";
    }

    private void ApplyRenames()
    {
        if (_videos.Count == 0 || _subtitles.Count == 0)
        {
            MessageBox.Show(this, "请先导入至少一个视频和一个字幕文件。", "没有可处理的项目",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var plans = BuildPlans();
        var readyPlans = plans.Where(plan => plan.CanRun).ToList();
        if (readyPlans.Count == 0)
        {
            MessageBox.Show(this, "当前没有可执行的重命名项目。请检查预览中的状态。", "无法执行",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var skipped = plans.Count - readyPlans.Count;
        var modeName = _createHardlink.Checked ? "创建硬链接副本，保留原字幕" : "直接重命名原字幕";
        var message = new StringBuilder()
            .AppendLine($"将执行 {readyPlans.Count} 个项目。")
            .AppendLine($"模式：{modeName}")
            .AppendLine($"覆盖已有目标：{(_overwriteExisting.Checked ? "是" : "否")}");

        if (_videos.Count != _subtitles.Count)
        {
            message.AppendLine($"注意：视频 {_videos.Count} 个，字幕 {_subtitles.Count} 个，将按预览中的匹配结果执行。");
        }

        if (skipped > 0)
        {
            message.AppendLine($"另外有 {skipped} 个项目因目标冲突、缺失或已同名而跳过。");
        }

        message.AppendLine().Append("确认继续？");
        var confirm = MessageBox.Show(this, message.ToString(), "确认应用更名",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
        if (confirm != DialogResult.OK)
        {
            return;
        }

        var successes = 0;
        var noOps = 0;
        var errors = new List<string>();

        foreach (var plan in readyPlans)
        {
            try
            {
                if (SamePath(plan.SourcePath, plan.TargetPath))
                {
                    noOps++;
                    continue;
                }

                if (_overwriteExisting.Checked && File.Exists(plan.TargetPath))
                {
                    File.Delete(plan.TargetPath);
                }

                Directory.CreateDirectory(Path.GetDirectoryName(plan.TargetPath)!);
                if (_createHardlink.Checked)
                {
                    CreateHardLinkOrThrow(plan.TargetPath, plan.SourcePath);
                }
                else
                {
                    File.Move(plan.SourcePath, plan.TargetPath, _overwriteExisting.Checked);
                }

                successes++;
            }
            catch (Exception ex)
            {
                errors.Add($"{Path.GetFileName(plan.SourcePath)} -> {Path.GetFileName(plan.TargetPath)}：{ex.Message}");
            }
        }

        if (!_createHardlink.Checked)
        {
            _subtitles.RemoveAll(item => !File.Exists(item.FullPath));
            foreach (var plan in readyPlans.Where(plan => File.Exists(plan.TargetPath)))
            {
                AddUnique(_subtitles, plan.TargetPath);
            }
        }

        _subtitles.Sort(CompareSubtitleByCanonicalName);
        RefreshAll($"完成：成功 {successes} 个，已同名 {noOps} 个，失败 {errors.Count} 个。");

        if (errors.Count > 0)
        {
            var detail = string.Join(Environment.NewLine, errors.Take(12));
            if (errors.Count > 12)
            {
                detail += Environment.NewLine + $"……还有 {errors.Count - 12} 个错误。";
            }

            MessageBox.Show(this, detail, "部分项目失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else
        {
            MessageBox.Show(this, $"处理完成。成功 {successes} 个，已同名 {noOps} 个。", "完成",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private List<RenamePlan> BuildPlans()
    {
        var plans = new List<RenamePlan>();
        var rows = BuildPreviewRows();
        plans.AddRange(rows
            .Where(row => row.Video is not null && row.Subtitle is not null)
            .Select(row => new RenamePlan(row.Subtitle!.FullPath, row.TargetPath, row.State is "就绪" or "已同名")));

        return plans;
    }

    private string GetEffectiveSuffix(FileItem video, FileItem subtitle)
    {
        var detected = CanonicalizeLanguageSuffix(DetectSuffix(video.FullPath, subtitle.FullPath));
        var custom = NormalizeSuffix(_suffixText.Text);

        return _suffixMode.SelectedIndex switch
        {
            1 => custom,
            2 => CombineSuffixes(custom, detected),
            3 => "",
            _ => string.IsNullOrEmpty(detected) ? custom : detected
        };
    }

    private string GetEffectiveSuffixForGrouping(FileItem subtitle)
    {
        var detected = CanonicalizeLanguageSuffix(DetectSuffixFromSubtitleName(subtitle.FullPath));
        var custom = NormalizeSuffix(_suffixText.Text);

        return _suffixMode.SelectedIndex switch
        {
            1 => custom,
            2 => string.IsNullOrEmpty(detected) ? custom : detected,
            3 => "",
            _ => string.IsNullOrEmpty(detected) ? custom : detected
        };
    }

    private static string CombineSuffixes(string prefixSuffix, string languageSuffix)
    {
        var first = NormalizeSuffix(prefixSuffix);
        var second = NormalizeSuffix(languageSuffix);
        if (first.Length == 0)
        {
            return second;
        }

        if (second.Length == 0)
        {
            return first;
        }

        return first + second;
    }

    private static string BuildTargetPath(FileItem video, FileItem subtitle, string suffix)
    {
        var videoStem = Path.GetFileNameWithoutExtension(video.FullPath);
        var extension = Path.GetExtension(subtitle.FullPath);
        var targetName = videoStem + suffix + extension;
        return Path.Combine(Path.GetDirectoryName(video.FullPath)!, targetName);
    }

    private static string DetectSuffix(string videoPath, string subtitlePath)
    {
        var videoStem = Path.GetFileNameWithoutExtension(videoPath);
        var subtitleStem = Path.GetFileNameWithoutExtension(subtitlePath);

        if (subtitleStem.Length > videoStem.Length &&
            subtitleStem.StartsWith(videoStem, StringComparison.OrdinalIgnoreCase))
        {
            return NormalizeSuffix(subtitleStem[videoStem.Length..]);
        }

        var tokens = ExtractTailSuffixTokens(subtitleStem);
        return tokens.Count == 0 ? "" : "." + string.Join(".", tokens);
    }

    private static string DetectSuffixFromSubtitleName(string subtitlePath)
    {
        var subtitleStem = Path.GetFileNameWithoutExtension(subtitlePath);
        var tokens = ExtractTailSuffixTokens(subtitleStem);
        return tokens.Count == 0 ? "" : "." + string.Join(".", tokens);
    }

    private string CanonicalizeDetectedSuffix(string suffix)
    {
        var normalized = NormalizeSuffix(suffix);
        if (normalized.Length == 0)
        {
            return "";
        }

        var outputTokens = new List<string>();
        foreach (var token in normalized.TrimStart('.').Split('.', StringSplitOptions.RemoveEmptyEntries))
        {
            outputTokens.Add(MapSuffixToken(token));
        }

        return outputTokens.Count == 0 ? "" : "." + string.Join(".", outputTokens);
    }

    private string CanonicalizeLanguageSuffix(string suffix)
    {
        var normalized = NormalizeSuffix(suffix);
        if (normalized.Length == 0)
        {
            return "";
        }

        var outputTokens = new List<string>();
        foreach (var token in normalized.TrimStart('.').Split('.', StringSplitOptions.RemoveEmptyEntries))
        {
            if (TryMapLanguageToken(token, out var mapped))
            {
                outputTokens.Add(mapped);
                continue;
            }

            var separators = new[] { '&', '+', '-', '_' };
            if (token.IndexOfAny(separators) < 0)
            {
                continue;
            }

            foreach (var part in token.Split(separators, StringSplitOptions.RemoveEmptyEntries))
            {
                if (TryMapLanguageToken(part, out var mappedPart))
                {
                    outputTokens.Add(mappedPart);
                }
            }
        }

        return outputTokens.Count == 0
            ? ""
            : "." + string.Join(".", outputTokens.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    private string MapSuffixToken(string token)
    {
        var clean = CleanToken(token);
        if (TryMapLanguageToken(clean, out var mapped))
        {
            return mapped;
        }

        var separators = new[] { '&', '+', '-', '_' };
        if (clean.IndexOfAny(separators) >= 0)
        {
            var parts = clean.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            var mappedParts = new List<string>();
            foreach (var part in parts)
            {
                if (!TryMapLanguageToken(part, out var mappedPart))
                {
                    mappedParts.Clear();
                    break;
                }

                mappedParts.Add(mappedPart);
            }

            if (mappedParts.Count > 0)
            {
                return string.Join(".", mappedParts);
            }
        }

        return clean;
    }

    private bool TryMapLanguageToken(string token, out string mapped)
    {
        var normalizedToken = NormalizeTokenForMatch(token);
        foreach (var rule in _languageRules)
        {
            var output = NormalizeSuffix(rule.OutputSuffix).TrimStart('.');
            if (output.Length > 0 && string.Equals(NormalizeTokenForMatch(output), normalizedToken, StringComparison.OrdinalIgnoreCase))
            {
                mapped = output;
                return true;
            }

            foreach (var alias in SplitAliases(rule.Aliases))
            {
                if (string.Equals(NormalizeTokenForMatch(alias), normalizedToken, StringComparison.OrdinalIgnoreCase))
                {
                    mapped = output.Length > 0 ? output : CleanToken(alias);
                    return true;
                }
            }
        }

        mapped = "";
        return false;
    }

    private static List<string> ExtractTailSuffixTokens(string stem)
    {
        var tokens = new List<string>();
        var text = stem.Trim();

        while (text.Length > 0)
        {
            var bracketMatch = Regex.Match(text, @"(?:[\[\(【｛]\s*(?<token>[^\]\)】｝]+?)\s*[\]\)】｝])\s*$");
            if (bracketMatch.Success)
            {
                var token = CleanToken(bracketMatch.Groups["token"].Value);
                if (!IsSuffixToken(token))
                {
                    break;
                }

                tokens.Insert(0, token);
                text = text[..bracketMatch.Index].TrimEnd('.', ' ', '_', '-');
                continue;
            }

            var dotMatch = Regex.Match(text, @"(?<prefix>.*?)[\._ ](?<token>[^._ ]+)\s*$");
            if (!dotMatch.Success)
            {
                break;
            }

            var dotToken = CleanToken(dotMatch.Groups["token"].Value);
            if (!IsSuffixToken(dotToken))
            {
                break;
            }

            tokens.Insert(0, dotToken);
            text = dotMatch.Groups["prefix"].Value.TrimEnd('.', ' ', '_', '-');
        }

        return tokens;
    }

    private static bool IsSuffixToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token) || IgnoredSuffixTokens.Contains(token))
        {
            return false;
        }

        if (LanguageTokens.Contains(token) || KnownGroupTokens.Contains(token))
        {
            return true;
        }

        var lower = token.ToLowerInvariant();
        if (lower.Contains('&') || lower.Contains('+') || lower.Contains('-') || lower.Contains('_'))
        {
            var parts = lower.Split(new[] { '&', '+', '-', '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 1 && parts.All(part => LanguageTokens.Contains(part) || KnownGroupTokens.Contains(part)))
            {
                return true;
            }
        }

        return Regex.IsMatch(token, @"^[A-Z][A-Z0-9]{1,11}$") && !IgnoredSuffixTokens.Contains(token);
    }

    private static string CleanToken(string token)
    {
        return token.Trim().Trim('.', ' ', '_', '-', '[', ']', '(', ')', '【', '】', '｛', '｝');
    }

    private static string NormalizeSuffix(string? suffix)
    {
        if (string.IsNullOrWhiteSpace(suffix))
        {
            return "";
        }

        var invalidChars = Path.GetInvalidFileNameChars();
        var cleaned = suffix.Trim().Replace('．', '.');
        cleaned = cleaned.Trim('.', ' ', '_', '-', '[', ']', '(', ')', '【', '】', '｛', '｝');
        if (cleaned.Length == 0)
        {
            return "";
        }

        cleaned = Regex.Replace(cleaned, @"\s+", ".");
        cleaned = new string(cleaned.Where(ch => !invalidChars.Contains(ch)).ToArray());
        cleaned = cleaned.Trim('.');
        return cleaned.Length == 0 ? "" : "." + cleaned;
    }

    private static void CreateHardLinkOrThrow(string targetPath, string sourcePath)
    {
        if (CreateHardLink(targetPath, sourcePath, IntPtr.Zero))
        {
            return;
        }

        var error = Marshal.GetLastWin32Error();
        throw new Win32Exception(error, "创建硬链接失败。请确认源字幕和目标目录在同一个 NTFS 分区上。");
    }

    private static bool SamePath(string left, string right)
    {
        return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
    }

    private static string SettingsPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SubtitleRenamer",
        "language-rules.json");

    private static List<LanguageRule> LoadLanguageRules()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath, Encoding.UTF8);
                var rules = JsonSerializer.Deserialize<List<LanguageRule>>(json);
                var sanitized = SanitizeRules(rules ?? new List<LanguageRule>());
                if (sanitized.Count > 0)
                {
                    return sanitized;
                }
            }
        }
        catch
        {
            // Broken settings should not prevent the tool from opening.
        }

        return DefaultLanguageRules();
    }

    private static void SaveLanguageRules(IReadOnlyList<LanguageRule> rules)
    {
        var directory = Path.GetDirectoryName(SettingsPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = JsonSerializer.Serialize(rules, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(SettingsPath, json, Encoding.UTF8);
    }

    private static List<LanguageRule> DefaultLanguageRules()
    {
        return new List<LanguageRule>
        {
            new("简体中文", ".chs", ".sc,.chs,.gb,.gbk,.zh-cn,.zh-hans,.zh-sg,.cn,.简,.简体,.简中,.JPSC,.chs&jpn"),
            new("繁体中文", ".cht", ".tc,.cht,.big5,.zh-tw,.zh-hant,.zh-hk,.繁,.繁体,.繁中,.JPTC,.cht&jpn"),
            new("日文", ".jpn", ".jpn,.jp,.ja,.japanese,.日,.日文,.日语"),
            new("英文", ".eng", ".eng,.en,.english,.英文,.英语")
        };
    }

    private static List<LanguageRule> CloneRules(IEnumerable<LanguageRule> rules)
    {
        return rules.Select(rule => new LanguageRule(rule.Language, rule.OutputSuffix, rule.Aliases)).ToList();
    }

    private static List<LanguageRule> SanitizeRules(IEnumerable<LanguageRule> rules)
    {
        var result = new List<LanguageRule>();
        foreach (var rule in rules)
        {
            var language = (rule.Language ?? "").Trim();
            var output = NormalizeSuffix(rule.OutputSuffix);
            var aliases = string.Join(",", SplitAliases(rule.Aliases).Select(alias => NormalizeSuffix(alias).TrimStart('.')).Where(alias => alias.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase));

            if (language.Length == 0 && output.Length == 0 && aliases.Length == 0)
            {
                continue;
            }

            if (language.Length == 0)
            {
                language = output.Length > 0 ? output.TrimStart('.') : "未命名语言";
            }

            result.Add(new LanguageRule(language, output, aliases));
        }

        return result;
    }

    private static void RegisterLanguageTokens(IEnumerable<LanguageRule> rules)
    {
        foreach (var rule in rules)
        {
            var output = NormalizeSuffix(rule.OutputSuffix).TrimStart('.');
            if (output.Length > 0)
            {
                LanguageTokens.Add(output);
            }

            foreach (var alias in SplitAliases(rule.Aliases))
            {
                var token = CleanToken(alias);
                if (token.Length > 0)
                {
                    LanguageTokens.Add(token);
                }
            }
        }
    }

    private static IEnumerable<string> SplitAliases(string? aliases)
    {
        return (aliases ?? "")
            .Split(new[] { ',', '，', ';', '；', '|', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(alias => CleanToken(alias))
            .Where(alias => alias.Length > 0);
    }

    private static string NormalizeTokenForMatch(string token)
    {
        return CleanToken(token).TrimStart('.').ToLowerInvariant();
    }

    private static string BuildImportMessage(ImportStats stats)
    {
        return $"导入：视频 {stats.Videos} 个，字幕 {stats.Subtitles} 个，排除 {stats.Ignored} 个，重复 {stats.Duplicates} 个。";
    }

    private enum ImportKind
    {
        Mixed,
        VideoOnly,
        SubtitleOnly
    }

    private sealed record FileItem(string FullPath, FileSource Source = FileSource.Local)
    {
        public string Name => Path.GetFileName(FullPath);
        public string DirectoryName => Path.GetDirectoryName(FullPath) ?? "";

        public static int CompareByName(FileItem left, FileItem right)
        {
            var nameCompare = StringComparer.CurrentCultureIgnoreCase.Compare(left.Name, right.Name);
            return nameCompare != 0
                ? nameCompare
                : StringComparer.CurrentCultureIgnoreCase.Compare(left.DirectoryName, right.DirectoryName);
        }
    }

    private sealed record CandidatePair(FileItem? Video, FileItem? Subtitle, string SlotSuffix = "");

    private sealed record PreviewRow(FileItem? Video, FileItem? Subtitle, string Suffix, string TargetPath, string State);

    private sealed record BlockPairResult(int BlockCount, List<CandidatePair> Pairs);

    private sealed class SubtitleGroup
    {
        public SubtitleGroup(string suffix)
        {
            Suffix = suffix;
        }

        public string Suffix { get; }
        public List<FileItem> Items { get; } = new();
    }

    private sealed record RenamePlan(string SourcePath, string TargetPath, bool CanRun);

    private sealed class ImportStats
    {
        public int Videos { get; set; }
        public int Subtitles { get; set; }
        public int Ignored { get; set; }
        public int Duplicates { get; set; }
    }
}
