# Fork Excel Merge Tool

Excel 三向合并与二向对比工具，兼容 **Fork** 的 Merge Tool 与 Diff Tool。Windows 桌面程序（.NET 8 / WPF），双击 `ExcelMergeFork.exe` 打开 Fluent 设置中心。

设置中心可与 [ExternalMergeTools](https://github.com/ZhouJun2303/ExternalMergeTools) 配合：对方作按类型分发器时，Excel 仍交给本工具。

---

## 一、功能概览

| 模式 | 作用 | 典型场景 |
|------|------|----------|
| **合并** | LOCAL + BASE + REMOTE → 一个合并结果 | Git 冲突时逐项选「本地」或「线上」 |
| **对比** | 比较两个 Excel，生成对比表 | Fork 里看本地 vs 线上 |

合并结果对新增 / 修改 / 冲突行做绿 / 黄 / 红标记，见 [颜色与场景说明.md](颜色与场景说明.md)。

---

## 二、获取与运行

- 需要 **Windows 64 位**。发布的 `ExcelMergeFork.exe` 是自包含包，**不用安装 Python 或 .NET**。
- 从 GitHub Releases 下载 `ExcelMergeFork-package.zip`，解压后双击 `install_fork_integration.bat` 注入 Fork。
- 也可只拿 `ExcelMergeFork.exe`：双击打开设置中心，在里面安装/移除 Fork 或 Git 注入。

已不再提供 lite / Python zip。旧的 `ExcelMergeFork-lite.exe` 与脚本入口已退役。

### 2.1 更新

设置中心里可检查 GitHub Releases。确认后才会下载并替换当前 exe。完整版资产名仍是 `ExcelMergeFork.exe`。

### 2.2 全局 Git 合并驱动（可选）

```text
install_git_integration.bat
```

会写入用户级 Git 配置，匹配常见 Excel 后缀。确认合并后只把结果写回 Git 的 `%A`，不执行 `git add`。卸载：

```text
uninstall_git_integration.bat
```

### 2.3 默认运行模式

双击 exe 打开设置中心。从 Fork、命令行或 Git 传入文件时按设置处理：

- **快速备份**：只备份输入文件
- **合并对比**：进入三向合并或二向对比（`.xlsx` / `.xltx`）
- **每次询问**：先弹窗选择。宏文件（`.xlsm` 等）请用快速备份，避免保存丢 VBA

---

## 三、在 Fork 里配置

### 3.1 一键注入（推荐）

先关闭 Fork，再执行 `install_fork_integration.bat`。移除用 `uninstall_fork_integration.bat`。也可在设置中心操作。

### 3.2 手动配置

- **Merge Tool**：Custom，路径指向本目录 `ExcelMergeFork.exe`，参数 `$LOCAL,$BASE,$REMOTE,$MERGED`
- **External Diff Tool**：路径同上，参数 `"$REMOTE" "$LOCAL"`

---

## 四、合并窗口

- **本地 (Local)**：当前分支；**线上 (Remote)**：被并入的分支。
- 差异列表虚拟滚动，冲突行可直接点「本地 / 线上」。
- 下方三栏显示 BASE / LOCAL / REMOTE 单元格。
- **生成合并结果** → 检查文件 → **确认并解决冲突**（普通 Fork 合并会 `git add`；Git driver 模式只写回 `%A`）。

备份目录默认为目标文件旁的 `MergeExcelBackup`，也可在设置里改根目录。

---

## 五、对比窗口

生成 `{原文件名}_compare.xlsx`。默认先在窗口里看差异，再导出工作簿。

---

## 六、命令行

```text
ExcelMergeFork.exe
ExcelMergeFork.exe <本地> <基准> <线上> <输出>
ExcelMergeFork.exe <线上REMOTE> <本地LOCAL>
ExcelMergeFork.exe --git-merge-driver <base> <current> <other> <repo-path>
ExcelMergeFork.exe --install-fork-integration
ExcelMergeFork.exe --install-git-integration
```

退出码：0 成功，1 参数/文件错误，2 处理异常。日志：`MergeExcelFork.log`。

---

## 七、开发者

需要 [.NET 8 SDK](https://dotnet.microsoft.com/download)。

```text
dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release
package.bat
```

`package.bat` 会测试并发布自包含 `ExcelMergeFork.exe` 到仓库根目录。`package.bat --dist --package-zip` 额外打 zip。

源码在 `src/`：`ExcelMergeFork.Core` 无 UI 引擎，`ExcelMergeFork.App` 为 WPF 界面，`ExcelMergeFork.Tests` 为 xUnit。夹具在 `TestData/`。
