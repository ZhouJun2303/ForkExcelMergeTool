# C# Fluent 重写实施说明

分支：`feat/csharp-fluent-rewrite`

选定方案：离开 Python，用 .NET 8 + WPF + WPF-UI + ClosedXML。界面为 Win11 设置风格；合并页上表下列、行内选本地/线上；WorkbookSession 一次读表。

已完成：Python / Tkinter / lite 已删除。根目录 `ExcelMergeFork.exe` 为 C# 自包含包，`package.bat` 只走 `dotnet publish`。

