namespace SubtitleRenamer;

internal sealed class LanguageSettingsDialog : Form
{
    private readonly DataGridView _grid;

    public LanguageSettingsDialog(List<LanguageRule> rules)
    {
        Text = "语言后缀设置";
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        Size = new Size(820, 460);
        MinimumSize = new Size(720, 390);
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
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        Controls.Add(root);

        root.Controls.Add(new Label
        {
            Dock = DockStyle.Fill,
            Text = "同一语言的多个写法会归为一组，并统一输出为你设置的后缀。例如 .sc 和 .chs 都可输出为 .chs。",
            ForeColor = Color.FromArgb(64, 72, 84),
            TextAlign = ContentAlignment.MiddleLeft
        }, 0, 0);

        _grid = BuildGrid();
        root.Controls.Add(_grid, 0, 1);

        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Padding = new Padding(0, 10, 0, 0),
            BackColor = BackColor
        };
        buttons.Controls.Add(CreateButton("取消", (_, _) => DialogResult = DialogResult.Cancel));
        buttons.Controls.Add(CreateButton("确定", (_, _) => Accept()));
        buttons.Controls.Add(CreateButton("恢复默认", (_, _) => LoadRules(DefaultRules())));
        buttons.Controls.Add(CreateButton("删除行", (_, _) => DeleteCurrentRow()));
        buttons.Controls.Add(CreateButton("新增行", (_, _) => _grid.Rows.Add("", ".", "")));
        root.Controls.Add(buttons, 0, 2);

        LoadRules(rules);
    }

    public List<LanguageRule> Rules { get; private set; } = new();

    private static DataGridView BuildGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            AllowUserToResizeRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
            EnableHeadersVisualStyles = false,
            GridColor = Color.FromArgb(231, 235, 241),
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };

        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(242, 245, 249);
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(45, 52, 64);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
        grid.ColumnHeadersHeight = 32;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(218, 234, 252);
        grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(20, 28, 38);
        grid.RowTemplate.Height = 28;

        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "语言名称", FillWeight = 20 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "输出后缀", FillWeight = 18 });
        grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "可识别写法（逗号分隔）", FillWeight = 62 });
        return grid;
    }

    private static Button CreateButton(string text, EventHandler handler)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = false,
            Size = new Size(text.Length > 3 ? 96 : 78, 32),
            Margin = new Padding(8, 0, 0, 0),
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

    private void LoadRules(IEnumerable<LanguageRule> rules)
    {
        _grid.Rows.Clear();
        foreach (var rule in rules)
        {
            _grid.Rows.Add(rule.Language, rule.OutputSuffix, rule.Aliases);
        }
    }

    private void DeleteCurrentRow()
    {
        if (_grid.CurrentRow is null || _grid.CurrentRow.IsNewRow)
        {
            return;
        }

        _grid.Rows.RemoveAt(_grid.CurrentRow.Index);
    }

    private void Accept()
    {
        _grid.EndEdit();
        var rules = new List<LanguageRule>();
        foreach (DataGridViewRow row in _grid.Rows)
        {
            if (row.IsNewRow)
            {
                continue;
            }

            var language = Convert.ToString(row.Cells[0].Value)?.Trim() ?? "";
            var output = Convert.ToString(row.Cells[1].Value)?.Trim() ?? "";
            var aliases = Convert.ToString(row.Cells[2].Value)?.Trim() ?? "";
            if (language.Length == 0 && output.Length == 0 && aliases.Length == 0)
            {
                continue;
            }

            rules.Add(new LanguageRule(language, output, aliases));
        }

        Rules = rules;
        DialogResult = DialogResult.OK;
    }

    private static List<LanguageRule> DefaultRules()
    {
        return new List<LanguageRule>
        {
            new("简体中文", ".chs", ".sc,.chs,.gb,.gbk,.zh-cn,.zh-hans,.zh-sg,.cn,.简,.简体,.简中"),
            new("繁体中文", ".cht", ".tc,.cht,.big5,.zh-tw,.zh-hant,.zh-hk,.繁,.繁体,.繁中"),
            new("日文", ".jpn", ".jpn,.jp,.ja,.japanese,.日,.日文,.日语"),
            new("英文", ".eng", ".eng,.en,.english,.英文,.英语")
        };
    }
}
