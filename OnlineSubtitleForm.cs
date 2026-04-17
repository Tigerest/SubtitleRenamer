namespace SubtitleRenamer;

internal sealed class OnlineSubtitleForm : Form
{
    private readonly OnlineSubtitleCache _cache;
    private readonly AcgripClient _client = new();
    private readonly Func<string> _defaultQueryProvider;
    private readonly Action<IReadOnlyList<string>> _importSubtitles;
    private readonly HashSet<string> _checkedSubtitlePaths = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _operationCts;
    private bool _allowClose;

    private TextBox _queryBox = null!;
    private Button _searchButton = null!;
    private Button _downloadButton = null!;
    private Button _importButton = null!;
    private Button _checkSelectedButton = null!;
    private Button _uncheckSelectedButton = null!;
    private Button _deleteSelectedButton = null!;
    private TreeView _resultTree = null!;
    private DataGridView _subtitleGrid = null!;
    private Label _statusLabel = null!;
    private Label _progressLabel = null!;
    private ProgressBar _progressBar = null!;
    private ContextMenuStrip _resultMenu = null!;
    private bool _subtitleDragSelecting;
    private int _subtitleDragStartRow = -1;

    private sealed record AttachmentNodeSelection(AcgripThreadFloor Floor, AcgripAttachment Attachment);

    public OnlineSubtitleForm(
        OnlineSubtitleCache cache,
        Func<string> defaultQueryProvider,
        Action<IReadOnlyList<string>> importSubtitles)
    {
        _cache = cache;
        _defaultQueryProvider = defaultQueryProvider;
        _importSubtitles = importSubtitles;

        BuildUi();
    }

    public void PrepareDefaultQuery()
    {
        if (!string.IsNullOrWhiteSpace(_queryBox.Text))
        {
            return;
        }

        _queryBox.Text = _defaultQueryProvider();
        _queryBox.SelectAll();
    }

    public void CloseForApplicationExit()
    {
        _allowClose = true;
        Close();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (!_allowClose && e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            Hide();
            return;
        }

        var cts = _operationCts;
        _operationCts = null;
        if (cts is not null)
        {
            try
            {
                cts.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }

            cts.Dispose();
        }

        _client.Dispose();
        base.OnFormClosing(e);
    }

    private void BuildUi()
    {
        Text = "在线查找字幕";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(1320, 860);
        MinimumSize = new Size(1080, 700);
        Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(246, 247, 249);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(12),
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        Controls.Add(root);

        root.Controls.Add(BuildSearchBar(), 0, 0);
        root.Controls.Add(BuildStatusBar(), 0, 1);
        root.Controls.Add(BuildSplitContent(), 0, 2);
    }

    private Control BuildSearchBar()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            BackColor = BackColor
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 330));

        panel.Controls.Add(new Label
        {
            Text = "剧名",
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);

        _queryBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 8, 8, 0)
        };
        _queryBox.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                _ = SearchAsync();
            }
        };
        panel.Controls.Add(_queryBox, 1, 0);

        _searchButton = CreateButton("搜索", (_, _) => _ = SearchAsync());
        _searchButton.Dock = DockStyle.Fill;
        panel.Controls.Add(_searchButton, 2, 0);

        var sourceLink = new LinkLabel
        {
            Text = "字幕来源：https://bbs.acgrip.com/",
            AutoSize = true,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            LinkColor = Color.FromArgb(30, 120, 212),
            ActiveLinkColor = Color.FromArgb(20, 88, 156),
            VisitedLinkColor = Color.FromArgb(30, 120, 212),
            Padding = new Padding(12, 10, 0, 0)
        };
        sourceLink.LinkClicked += (_, _) => OpenUrl("https://bbs.acgrip.com/");
        panel.Controls.Add(sourceLink, 3, 0);

        return panel;
    }

    private Control BuildStatusBar()
    {
        _statusLabel = new Label
        {
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(86, 96, 112),
            TextAlign = ContentAlignment.MiddleLeft,
            Text = "搜索会复用本次运行缓存，并限制请求频率，避免给 ACGRIP 带来额外压力。"
        };
        return _statusLabel;
    }

    private Control BuildSplitContent()
    {
        var splitterInitialized = false;
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            SplitterWidth = 7,
            BackColor = BackColor
        };
        split.SizeChanged += (_, _) =>
        {
            if (splitterInitialized)
            {
                return;
            }

            const int preferredTopHeight = 390;
            const int preferredMinPanelHeight = 180;
            var maxDistance = split.Height - preferredMinPanelHeight - split.SplitterWidth;
            if (maxDistance <= preferredMinPanelHeight)
            {
                return;
            }

            split.SplitterDistance = Math.Clamp(preferredTopHeight, preferredMinPanelHeight, maxDistance);
            splitterInitialized = true;
        };
        split.Panel1.Controls.Add(BuildResultArea());
        split.Panel2.Controls.Add(BuildSubtitlePanel());
        return split;
    }

    private Control BuildResultArea()
    {
        var group = new GroupBox
        {
            Text = "搜索结果：悬浮查看详情，右键选择下载或打开帖子",
            Dock = DockStyle.Fill,
            Padding = new Padding(10),
            BackColor = Color.White
        };

        _resultTree = new TreeView
        {
            Dock = DockStyle.Fill,
            HideSelection = false,
            BorderStyle = BorderStyle.None,
            ShowNodeToolTips = true,
            FullRowSelect = true
        };
        _resultMenu = new ContextMenuStrip();
        _resultMenu.Items.Add("下载并解压", null, (_, _) => _ = DownloadSelectedAsync());
        _resultMenu.Items.Add("打开帖子网页", null, (_, _) => OpenSelectedThread());
        _resultTree.ContextMenuStrip = _resultMenu;
        _resultTree.AfterSelect += (_, _) => _downloadButton.Enabled = GetSelectedDownloadItems().Count > 0;
        _resultTree.NodeMouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Right)
            {
                _resultTree.SelectedNode = e.Node;
            }
        };
        group.Controls.Add(_resultTree);
        return group;
    }

    private Control BuildSubtitlePanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = BackColor
        };
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        panel.Controls.Add(BuildDownloadBar(), 0, 0);
        panel.Controls.Add(BuildSubtitleArea(), 0, 1);
        return panel;
    }

    private Control BuildDownloadBar()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            BackColor = BackColor
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 136));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));

        _downloadButton = CreateButton("下载并解压", (_, _) => _ = DownloadSelectedAsync());
        _downloadButton.Enabled = false;
        _downloadButton.Dock = DockStyle.Fill;
        panel.Controls.Add(_downloadButton, 0, 0);

        _progressBar = new ProgressBar
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(12, 13, 8, 12),
            Style = ProgressBarStyle.Continuous
        };
        panel.Controls.Add(_progressBar, 1, 0);

        _progressLabel = new Label
        {
            Dock = DockStyle.Fill,
            Text = "等待下载",
            TextAlign = ContentAlignment.MiddleLeft,
            ForeColor = Color.FromArgb(86, 96, 112)
        };
        panel.Controls.Add(_progressLabel, 2, 0);

        var hint = new Label
        {
            Dock = DockStyle.Fill,
            Text = "可拖动上方分隔线",
            TextAlign = ContentAlignment.MiddleRight,
            ForeColor = Color.FromArgb(120, 128, 140)
        };
        panel.Controls.Add(hint, 3, 0);

        return panel;
    }

    private Control BuildSubtitleArea()
    {
        var group = new GroupBox
        {
            Text = "已解压字幕：按压缩包内文件夹分组，支持 Ctrl/Shift 多选后批量勾选",
            Dock = DockStyle.Fill,
            Padding = new Padding(10),
            BackColor = Color.White
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

        _subtitleGrid = new DataGridView
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
            MultiSelect = true,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };
        _subtitleGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(242, 245, 249);
        _subtitleGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(45, 52, 64);
        _subtitleGrid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
        _subtitleGrid.ColumnHeadersHeight = 32;
        _subtitleGrid.RowTemplate.Height = 28;
        _subtitleGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "导入", FillWeight = 8 });
        _subtitleGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "文件夹", FillWeight = 18, ReadOnly = true });
        _subtitleGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "字幕文件", FillWeight = 30, ReadOnly = true });
        _subtitleGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "来源帖子/楼层", FillWeight = 30, ReadOnly = true });
        _subtitleGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "附件", FillWeight = 18, ReadOnly = true });
        _subtitleGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", FillWeight = 10, ReadOnly = true });
        _subtitleGrid.CellClick += (_, e) =>
        {
            if (e.RowIndex >= 0)
            {
                ToggleRowCheck(_subtitleGrid.Rows[e.RowIndex]);
            }
        };
        _subtitleGrid.MouseDown += HandleSubtitleGridMouseDown;
        _subtitleGrid.MouseMove += HandleSubtitleGridMouseMove;
        _subtitleGrid.MouseUp += (_, _) => _subtitleDragSelecting = false;
        layout.Controls.Add(_subtitleGrid, 0, 0);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Padding = new Padding(0, 8, 0, 0),
            BackColor = Color.White
        };
        _importButton = CreateButton("导入主窗口", (_, _) => ImportCheckedSubtitles());
        _importButton.BackColor = Color.FromArgb(30, 120, 212);
        _importButton.ForeColor = Color.White;
        _importButton.FlatAppearance.BorderColor = Color.FromArgb(30, 120, 212);
        buttons.Controls.Add(_importButton);
        buttons.Controls.Add(CreateButton("全选", (_, _) => SetAllSubtitleChecks(true)));
        buttons.Controls.Add(CreateButton("全不选", (_, _) => SetAllSubtitleChecks(false)));
        _checkSelectedButton = CreateButton("勾选", (_, _) => SetSelectedSubtitleChecks(true));
        _uncheckSelectedButton = CreateButton("取消勾选", (_, _) => SetSelectedSubtitleChecks(false));
        _deleteSelectedButton = CreateButton("删除选中", (_, _) => DeleteSelectedSubtitles());
        buttons.Controls.Add(_uncheckSelectedButton);
        buttons.Controls.Add(_checkSelectedButton);
        buttons.Controls.Add(_deleteSelectedButton);
        layout.Controls.Add(buttons, 0, 1);

        group.Controls.Add(layout);
        RefreshSubtitleGrid();
        return group;
    }

    private static Button CreateButton(string text, EventHandler handler)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = false,
            Size = new Size(text.Length >= 4 ? 124 : 82, 34),
            Margin = new Padding(8, 3, 0, 5),
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

    private async Task SearchAsync()
    {
        var query = _queryBox.Text.Trim();
        if (query.Length == 0)
        {
            MessageBox.Show(this, "请输入要搜索的剧名。", "缺少搜索词", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _operationCts?.Cancel();
        _operationCts = new CancellationTokenSource();
        SetBusy(true, useMarquee: true);

        try
        {
            if (_cache.TryGetSearch(query, out var cached))
            {
                FillResultTree(cached.Floors);
                _statusLabel.Text = $"已从本次缓存读取：{query}，共 {cached.Floors.Count} 个含字幕附件的楼层。";
                return;
            }

            _statusLabel.Text = "正在搜索 ACGRIP...";
            var results = await _client.SearchAsync(query, maxResults: 6, _operationCts.Token);
            var floors = new List<AcgripThreadFloor>();
            var checkedThreads = 0;
            _resultTree.Nodes.Clear();

            foreach (var result in results.Take(5))
            {
                checkedThreads++;
                _statusLabel.Text = $"正在读取帖子 {checkedThreads}/{Math.Min(results.Count, 5)}：{result.Title}";
                var threadFloors = await _client.FetchSubtitleFloorsAsync(result, maxPages: 2, _operationCts.Token);
                floors.AddRange(threadFloors);
                AppendResultFloors(threadFloors);
                _statusLabel.Text = $"已显示 {floors.Count} 个含字幕附件的楼层，继续搜索中...";
            }

            var snapshot = new AcgripSearchSnapshot(query, DateTime.Now, floors);
            _cache.SaveSearch(snapshot);
            _statusLabel.Text = floors.Count == 0
                ? "没有找到带字幕附件的楼层，可以换个剧名再试。"
                : $"搜索完成：找到 {floors.Count} 个含字幕附件的楼层。";
        }
        catch (OperationCanceledException)
        {
            _statusLabel.Text = "操作已取消。";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "搜索失败。";
            MessageBox.Show(this, ex.Message, "在线搜索失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private async Task DownloadSelectedAsync()
    {
        var attachments = GetSelectedDownloadItems();
        if (attachments.Count == 0)
        {
            return;
        }

        _operationCts?.Cancel();
        _operationCts = new CancellationTokenSource();
        SetBusy(true);

        try
        {
            var added = 0;
            for (var i = 0; i < attachments.Count; i++)
            {
                var (floor, attachment) = attachments[i];
                _statusLabel.Text = $"正在处理 {i + 1}/{attachments.Count}：{attachment.Name}";
                ResetProgress();

                string downloadedPath;
                if (!_cache.TryGetDownloadedAttachment(attachment, out downloadedPath!))
                {
                    var targetPath = _cache.GetDownloadPath(attachment, attachment.Name);
                    var progress = new Progress<AcgripDownloadProgress>(UpdateDownloadProgress);
                    downloadedPath = await _client.DownloadAttachmentAsync(
                        attachment,
                        targetPath,
                        floor.Thread.Url,
                        progress,
                        _operationCts.Token);
                    _cache.SaveDownloadedAttachment(attachment, downloadedPath);
                }
                else
                {
                    _progressLabel.Text = "使用本次缓存";
                    _progressBar.Value = 100;
                }

                var extractDir = _cache.GetExtractDirectory(attachment);
                _progressLabel.Text = "正在解压...";
                var subtitles = await ArchiveExtractor.ExtractSubtitlesAsync(downloadedPath, extractDir, _operationCts.Token);
                added += _cache.AddExtractedSubtitles(subtitles, extractDir, floor, attachment);
            }

            RefreshSubtitleGrid();
            _progressBar.Value = 100;
            _progressLabel.Text = "处理完成";
            _statusLabel.Text = added == 0
                ? "已处理选中楼层，没有新增字幕。"
                : $"已新增 {added} 个解压字幕，可在下方勾选后导入主窗口。";
        }
        catch (OperationCanceledException)
        {
            _statusLabel.Text = "操作已取消。";
        }
        catch (Exception ex)
        {
            _statusLabel.Text = "下载或解压失败。";
            MessageBox.Show(this, ex.Message, "下载或解压失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void FillResultTree(IReadOnlyList<AcgripThreadFloor> floors)
    {
        _resultTree.BeginUpdate();
        _resultTree.Nodes.Clear();
        foreach (var group in floors.GroupBy(floor => floor.Thread.Tid))
        {
            _resultTree.Nodes.Add(BuildThreadNode(group.ToList()));
        }

        _resultTree.EndUpdate();
        _downloadButton.Enabled = GetSelectedDownloadItems().Count > 0;
    }

    private void AppendResultFloors(IReadOnlyList<AcgripThreadFloor> floors)
    {
        if (floors.Count == 0)
        {
            return;
        }

        _resultTree.BeginUpdate();
        foreach (var group in floors.GroupBy(floor => floor.Thread.Tid))
        {
            _resultTree.Nodes.Add(BuildThreadNode(group.ToList()));
        }

        _resultTree.EndUpdate();
        _downloadButton.Enabled = GetSelectedDownloadItems().Count > 0;
    }

    private static TreeNode BuildThreadNode(IReadOnlyList<AcgripThreadFloor> floors)
    {
        var thread = floors[0].Thread;
        var threadNode = new TreeNode($"{thread.Title}    [{thread.ForumName}]    匹配 {thread.Score}")
        {
            Tag = thread,
            ToolTipText = $"{thread.Url}\r\n{thread.Snippet}"
        };

        foreach (var floor in floors)
        {
            var attachmentText = string.Join("  |  ", floor.Attachments.Select(item => item.Name));
            var meta = string.Join(" / ", new[] { floor.FloorLabel, floor.Author, floor.PostedAt }.Where(item => !string.IsNullOrWhiteSpace(item)));
            var node = new TreeNode($"{meta}    附件 {floor.Attachments.Count} 个")
            {
                Tag = floor,
                ToolTipText = string.Join("\r\n", new[] { attachmentText, floor.Excerpt }.Where(item => !string.IsNullOrWhiteSpace(item)))
            };

            foreach (var attachment in floor.Attachments)
            {
                node.Nodes.Add(new TreeNode($"    {attachment.Name}")
                {
                    Tag = new AttachmentNodeSelection(floor, attachment),
                    ToolTipText = $"{attachment.Name}\r\n{floor.Thread.Url}"
                });
            }

            threadNode.Nodes.Add(node);
        }

        threadNode.Expand();
        return threadNode;
    }

    private List<AcgripThreadFloor> GetSelectedFloors()
    {
        if (_resultTree.SelectedNode?.Tag is AttachmentNodeSelection selected)
        {
            return new List<AcgripThreadFloor> { selected.Floor };
        }

        if (_resultTree.SelectedNode?.Tag is AcgripThreadFloor floor)
        {
            return new List<AcgripThreadFloor> { floor };
        }

        if (_resultTree.SelectedNode?.Tag is AcgripThreadResult)
        {
            return _resultTree.SelectedNode.Nodes
                .Cast<TreeNode>()
                .Select(node => node.Tag)
                .OfType<AcgripThreadFloor>()
                .ToList();
        }

        return new List<AcgripThreadFloor>();
    }

    private List<(AcgripThreadFloor Floor, AcgripAttachment Attachment)> GetSelectedDownloadItems()
    {
        return _resultTree.SelectedNode?.Tag switch
        {
            AttachmentNodeSelection selected => new List<(AcgripThreadFloor, AcgripAttachment)>
            {
                (selected.Floor, selected.Attachment)
            },
            AcgripThreadFloor floor => floor.Attachments.Select(attachment => (floor, attachment)).ToList(),
            AcgripThreadResult => GetSelectedFloors()
                .SelectMany(floor => floor.Attachments.Select(attachment => (floor, attachment)))
                .ToList(),
            _ => new List<(AcgripThreadFloor, AcgripAttachment)>()
        };
    }

    private void OpenSelectedThread()
    {
        switch (_resultTree.SelectedNode?.Tag)
        {
            case AttachmentNodeSelection selected:
                OpenUrl(selected.Floor.Thread.Url);
                break;
            case AcgripThreadFloor floor:
                OpenUrl(floor.Thread.Url);
                break;
            case AcgripThreadResult thread:
                OpenUrl(thread.Url);
                break;
        }
    }

    private void RefreshSubtitleGrid()
    {
        _subtitleGrid.Rows.Clear();
        foreach (var group in _cache.Subtitles.GroupBy(item => $"{item.AttachmentName}::{item.FolderLabel}"))
        {
            var items = group.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
            var first = items[0];
            var headerChecked = items.Any(item => !item.Imported && _checkedSubtitlePaths.Contains(item.FullPath));
            var headerIndex = _subtitleGrid.Rows.Add(headerChecked, first.FolderLabel, $"{first.AttachmentName} / {first.FolderLabel}", first.ThreadTitle, "", "");
            var header = _subtitleGrid.Rows[headerIndex];
            header.Tag = items;
            header.DefaultCellStyle.BackColor = Color.FromArgb(226, 238, 233);
            header.DefaultCellStyle.ForeColor = Color.FromArgb(25, 88, 73);
            header.DefaultCellStyle.Font = new Font(_subtitleGrid.Font, FontStyle.Bold);

            foreach (var item in items)
            {
                var checkedValue = !item.Imported && _checkedSubtitlePaths.Contains(item.FullPath);
                var rowIndex = _subtitleGrid.Rows.Add(
                    checkedValue,
                    "",
                    item.Name,
                    $"{item.ThreadTitle} / {item.FloorLabel}",
                    item.AttachmentName,
                    item.Imported ? "已导入" : "待导入");
                _subtitleGrid.Rows[rowIndex].Tag = item;
                if (item.Imported)
                {
                    _subtitleGrid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.FromArgb(128, 128, 128);
                }
            }
        }
    }

    private void SetAllSubtitleChecks(bool value)
    {
        foreach (DataGridViewRow row in _subtitleGrid.Rows)
        {
            SetRowCheck(row, value);
        }
    }

    private void SetSelectedSubtitleChecks(bool value)
    {
        foreach (DataGridViewRow row in _subtitleGrid.SelectedRows)
        {
            SetRowCheck(row, value);
        }
    }

    private void HandleSubtitleGridMouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button != MouseButtons.Left || ModifierKeys != Keys.None)
        {
            return;
        }

        var hit = _subtitleGrid.HitTest(e.X, e.Y);
        if (hit.RowIndex < 0 || hit.ColumnIndex == 0)
        {
            return;
        }

        _subtitleDragSelecting = true;
        _subtitleDragStartRow = hit.RowIndex;
        SelectSubtitleRowRange(_subtitleDragStartRow, hit.RowIndex);
    }

    private void HandleSubtitleGridMouseMove(object? sender, MouseEventArgs e)
    {
        if (!_subtitleDragSelecting || _subtitleDragStartRow < 0)
        {
            return;
        }

        var hit = _subtitleGrid.HitTest(e.X, e.Y);
        if (hit.RowIndex >= 0)
        {
            SelectSubtitleRowRange(_subtitleDragStartRow, hit.RowIndex);
        }
    }

    private void SelectSubtitleRowRange(int startRow, int endRow)
    {
        var first = Math.Min(startRow, endRow);
        var last = Math.Max(startRow, endRow);
        _subtitleGrid.ClearSelection();
        for (var i = first; i <= last && i < _subtitleGrid.Rows.Count; i++)
        {
            _subtitleGrid.Rows[i].Selected = true;
        }
    }

    private void ToggleRowCheck(DataGridViewRow row)
    {
        if (row.Tag is List<OnlineSubtitleItem> groupItems)
        {
            var anyUnchecked = groupItems.Any(item => !item.Imported && !_checkedSubtitlePaths.Contains(item.FullPath));
            SetFolderGroupCheck(row, anyUnchecked);
            return;
        }

        if (row.Tag is not OnlineSubtitleItem item || item.Imported)
        {
            return;
        }

        SetRowCheck(row, !Convert.ToBoolean(row.Cells[0].Value));
    }

    private void SetRowCheck(DataGridViewRow row, bool value)
    {
        if (row.Tag is List<OnlineSubtitleItem> groupItems)
        {
            SetFolderGroupCheck(row, value);
            return;
        }

        if (row.Tag is not OnlineSubtitleItem item || item.Imported)
        {
            return;
        }

        row.Cells[0].Value = value;
        if (value)
        {
            _checkedSubtitlePaths.Add(item.FullPath);
        }
        else
        {
            _checkedSubtitlePaths.Remove(item.FullPath);
        }
    }

    private void SetFolderGroupCheck(DataGridViewRow headerRow, bool value)
    {
        if (headerRow.Tag is not List<OnlineSubtitleItem> groupItems)
        {
            return;
        }

        headerRow.Cells[0].Value = value;
        foreach (var item in groupItems.Where(item => !item.Imported))
        {
            if (value)
            {
                _checkedSubtitlePaths.Add(item.FullPath);
            }
            else
            {
                _checkedSubtitlePaths.Remove(item.FullPath);
            }
        }

        foreach (DataGridViewRow row in _subtitleGrid.Rows)
        {
            if (row.Tag is OnlineSubtitleItem item && groupItems.Contains(item))
            {
                row.Cells[0].Value = value && !item.Imported;
            }
        }
    }

    private void ImportCheckedSubtitles()
    {
        _subtitleGrid.EndEdit();
        var selected = new List<OnlineSubtitleItem>();
        foreach (DataGridViewRow row in _subtitleGrid.Rows)
        {
            if (row.Tag is not OnlineSubtitleItem item || item.Imported)
            {
                continue;
            }

            var checkedValue = Convert.ToBoolean(row.Cells[0].Value);
            if (checkedValue)
            {
                selected.Add(item);
                _checkedSubtitlePaths.Add(item.FullPath);
            }
            else
            {
                _checkedSubtitlePaths.Remove(item.FullPath);
            }
        }

        if (selected.Count == 0)
        {
            MessageBox.Show(this, "请先勾选要导入主窗口的字幕。", "没有选中字幕", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _importSubtitles(selected.Select(item => item.FullPath).ToList());
        foreach (var item in selected)
        {
            item.Imported = true;
            _checkedSubtitlePaths.Remove(item.FullPath);
        }

        RefreshSubtitleGrid();
        _statusLabel.Text = $"已导入 {selected.Count} 个字幕到主窗口。";
    }

    public void ResetImportedState(IEnumerable<string>? paths = null)
    {
        var pathSet = paths?.Select(Path.GetFullPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var changed = false;
        foreach (var item in _cache.Subtitles)
        {
            if (pathSet is not null && !pathSet.Contains(Path.GetFullPath(item.FullPath)))
            {
                continue;
            }

            if (item.Imported)
            {
                item.Imported = false;
                changed = true;
            }
        }

        if (!changed)
        {
            return;
        }

        _cache.Save();
        RefreshSubtitleGrid();
    }

    private void DeleteSelectedSubtitles()
    {
        var selected = GetSelectedSubtitleItems();
        if (selected.Count == 0)
        {
            MessageBox.Show(this, "请先选择要删除的已解压字幕。", "没有选中字幕", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var confirm = MessageBox.Show(this, $"删除选中的 {selected.Count} 个已解压字幕？", "确认删除",
            MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
        if (confirm != DialogResult.OK)
        {
            return;
        }

        var paths = selected.Select(item => item.FullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var removed = _cache.RemoveSubtitles(paths, deleteFiles: true);
        foreach (var path in paths)
        {
            _checkedSubtitlePaths.Remove(path);
        }

        RefreshSubtitleGrid();
        _statusLabel.Text = $"已删除 {removed} 个已解压字幕。";
    }

    private List<OnlineSubtitleItem> GetSelectedSubtitleItems()
    {
        var result = new List<OnlineSubtitleItem>();
        foreach (DataGridViewRow row in _subtitleGrid.SelectedRows)
        {
            if (row.Tag is OnlineSubtitleItem item)
            {
                result.Add(item);
            }
            else if (row.Tag is List<OnlineSubtitleItem> groupItems)
            {
                result.AddRange(groupItems);
            }
        }

        return result.DistinctBy(item => item.FullPath).ToList();
    }

    private void UpdateDownloadProgress(AcgripDownloadProgress progress)
    {
        if (progress.TotalBytes is > 0)
        {
            var percent = Math.Clamp((int)(progress.BytesReceived * 100 / progress.TotalBytes.Value), 0, 100);
            _progressBar.Style = ProgressBarStyle.Continuous;
            _progressBar.Value = percent;
            _progressLabel.Text = $"{FormatBytes(progress.BytesReceived)} / {FormatBytes(progress.TotalBytes.Value)}  {FormatBytes((long)progress.BytesPerSecond)}/s";
        }
        else
        {
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressLabel.Text = $"{FormatBytes(progress.BytesReceived)}  {FormatBytes((long)progress.BytesPerSecond)}/s";
        }
    }

    private void ResetProgress()
    {
        _progressBar.Style = ProgressBarStyle.Continuous;
        _progressBar.Value = 0;
        _progressLabel.Text = "准备下载";
    }

    private void SetBusy(bool busy, bool useMarquee = false)
    {
        _searchButton.Enabled = !busy;
        _downloadButton.Enabled = !busy && GetSelectedDownloadItems().Count > 0;
        _importButton.Enabled = !busy;
        _checkSelectedButton.Enabled = !busy;
        _uncheckSelectedButton.Enabled = !busy;
        _deleteSelectedButton.Enabled = !busy;
        if (busy)
        {
            Cursor = Cursors.WaitCursor;
            if (useMarquee)
            {
                _progressBar.Style = ProgressBarStyle.Marquee;
                _progressLabel.Text = "正在加载...";
            }
        }
        else
        {
            Cursor = Cursors.Default;
            UseWaitCursor = false;
            var wasLoading = _progressBar.Style == ProgressBarStyle.Marquee || _progressLabel.Text == "正在加载...";
            _progressBar.Style = ProgressBarStyle.Continuous;
            if (wasLoading || _progressBar.Value < 100)
            {
                _progressBar.Value = 0;
                _progressLabel.Text = "等待下载";
            }
        }
    }

    private static string FormatBytes(long bytes)
    {
        string[] units = { "B", "KB", "MB", "GB" };
        var value = (double)Math.Max(0, bytes);
        var unit = 0;
        while (value >= 1024 && unit < units.Length - 1)
        {
            value /= 1024;
            unit++;
        }

        return unit == 0 ? $"{value:0} {units[unit]}" : $"{value:0.0} {units[unit]}";
    }

    private static void OpenUrl(string url)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }
}
