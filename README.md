# SubtitleRenamer

一个轻量的 Windows 外挂字幕重命名工具。它可以把视频和字幕按预览关系匹配，并将字幕重命名为播放器容易自动识别的同名外挂字幕。

## 功能

- 拖入视频、字幕或文件夹，自动排除非视频和非常规字幕文件。
- 视频与字幕导入后自动排序，视频列表可手动上移、下移。
- 支持多语言字幕批量匹配，例如简体、繁体、日文字幕一次处理。
- 支持语言后缀规则设置，例如将 `.sc`、`.chs`、`.zh-cn` 统一识别为 `.chs`。
- 支持默认后缀策略：
  - 识别原语言后缀
  - 统一使用默认后缀
  - 默认后缀 + 语言后缀
  - 不添加后缀
- 预览窗口支持拖动空行来调整字幕缺失位置。
- 支持直接重命名原字幕，或创建硬链接副本并保留原字幕。
- 支持覆盖已有同名字幕。

## 构建

需要 Windows 和 .NET SDK 7.0 或更新版本。

```powershell
dotnet build -c Release
```

发布为单文件 exe：

```powershell
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained false -p:DebugType=None -p:DebugSymbols=false -o publish
```

发布后的文件位于：

```text
publish/SubtitleRenamer.exe
```

## 首次运行提示

如果 Windows 提示“无法确认是谁创建了此文件”，通常是因为 exe 未签名或带有下载来源标记。自己使用时可以右键文件属性解除锁定，或运行：

```powershell
Unblock-File -LiteralPath ".\SubtitleRenamer.exe"
```

## 配置

语言后缀设置会保存在：

```text
%AppData%\SubtitleRenamer\language-rules.json
```

## 许可证

MIT License
