# Fork Excel Merge Tool

Excel 三向合并与二向对比工具，兼容 Fork 客户端的 Merge Tool 与 Diff Tool。支持 N 个 Sheet，合并后自动备份三方 Excel。

## 功能

- **合并模式**：三向合并（LOCAL + BASE + REMOTE → MERGED），冲突取 LOCAL 并标红
- **对比模式**：二向 diff（LOCAL vs REMOTE），生成对比 Excel 并自动打开

颜色与各场景说明见 [颜色与场景说明.md](颜色与场景说明.md)。

## 系统要求

- **Windows**（推荐）：使用 `ExcelMergeFork.exe`，无需安装 Python
- **Python 运行**：需 Python 3.7+，执行 `pip install openpyxl`

## 获取方式

复制整个 `ForkExcelMergeTool` 文件夹到任意路径即可使用。

最小分发包：`ExcelMergeFork.exe` + `README.md`

## Fork 配置

### Merge Tool（解决冲突）

1. Fork → Preferences → Integration → Merge Tool
2. Merger：选择 **Custom**
3. Merger Path：`ExcelMergeFork.exe` 的完整路径
4. Arguments：`$LOCAL $BASE $REMOTE $MERGED` 或 `$LOCAL,$BASE,$REMOTE,$MERGED`（Fork 可能用逗号合并为单参数，已支持）

### Diff Tool（查看差异）

1. Fork → Preferences → Integration → External Diff Tool
2. Diff Tool：选择 **Custom**
3. Diff Tool Path：同上 `ExcelMergeFork.exe` 路径
4. Arguments：`$LOCAL $REMOTE` 或 `$LOCAL,$REMOTE`（Fork 可能用逗号合并为单参数，已支持）

## 独立用法（命令行）

**合并**：

```
ExcelMergeFork.exe <local> <base> <remote> <merged>
```

**对比**：

```
ExcelMergeFork.exe <fileA> <fileB>
```

生成 `{fileA 同名}_compare.xlsx` 并自动打开。

## 打包 exe

**一键打包**：双击 `一键打包.bat`，自动安装依赖、打包、并运行快速验证。

或手动打包：

```
pip install pyinstaller
build_exe.bat
```

## 测试

- **一键测试合并**：双击 `run_quick_test.bat`（自动生成测试数据 → 合并 → 打开输出目录）
- **手动测试合并**：`run_merge.bat TestData\local.xlsx TestData\base.xlsx TestData\remote.xlsx TestData\_output\merged.xlsx`
- **手动测试对比**：`run_compare.bat TestData\local.xlsx TestData\remote.xlsx`

## 备份目录

合并完成后，LOCAL、REMOTE、MERGED 会复制到 **MERGED 文件所在目录**（即仓库内冲突文件所在目录）：

`{MERGED 文件名}_local.xlsx`、`_remote.xlsx`、`_merged.xlsx`

Fork 调用时 LOCAL/BASE/REMOTE 在临时目录，备份会写入仓库目录，便于查找。

## 退出码

- 0：成功
- 1：参数或文件错误
- 2：合并/对比异常

## 日志与输出

- **日志**：与 exe 同目录下的 `MergeExcelFork.log`，异常会写入完整堆栈，闪退时请查看此文件
- **对比输出**：`{第一个文件所在目录}/{第一个文件名}_compare.xlsx`，成功后会尝试用 Excel 打开
- **无图形化面板**：对比模式生成 Excel 文件并调用系统默认程序打开，不提供独立 GUI 窗口。若需在 Fork 中查看差异，请手动打开生成的 `*_compare.xlsx`
