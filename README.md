# Fork Excel Merge Tool

Excel 三向合并与二向对比工具，兼容 **Fork** 客户端的 Merge Tool 与 Diff Tool。支持多 Sheet，合并后自动备份，冲突可在图形界面里逐项选「本地」或「线上」。

---

## 一、功能概览

| 模式 | 作用 | 典型场景 |
|------|------|----------|
| **合并** | LOCAL + BASE + REMOTE → 一个合并结果文件 | Git 合并冲突时，用本工具选每处冲突保留哪边 |
| **对比** | 比较两个 Excel 的差异，生成对比表 | 在 Fork 里查看某文件的「本地 vs 线上」差异 |

- 合并结果会对**新增 / 修改 / 冲突**行做颜色标记（绿 / 黄 / 红），详见 [颜色与场景说明.md](颜色与场景说明.md)。

---

## 二、系统要求与获取方式

- **Windows（推荐）**：使用打包好的 `ExcelMergeFork.exe`，**无需安装 Python**。
- **Python 运行**：需 Python 3.7+，执行 `pip install openpyxl` 后即可用脚本。

**推荐分发方式**：从 GitHub Releases 下载最新版 `ExcelMergeFork.exe`。
**最小分发包**：只需 `ExcelMergeFork.exe` + 本 `README.md`（对方无需其它文件即可在 Fork 里使用）。Fork 里配置的是 exe 的完整路径，后续应用内更新会原地替换这个 exe，Fork 配置不用改。

### 2.1 更新方式

工具会在合并/对比窗口后台检查 GitHub Releases。检查到新版本时，界面上的「检查更新」按钮会变成「有新版本 vX.Y」。

- 不会自动下载或替换，必须由用户点击按钮并确认后才更新。
- 检查和下载期间按钮会禁用，窗口会显示更新状态；下载时有进度条，无法读取总大小时显示滚动进度。
- 更新时会下载最新版 `ExcelMergeFork.exe`，如果 Release 里带有 `ExcelMergeFork.exe.sha256` 会自动校验。
- Windows 正在运行的 exe 不能直接覆盖，所以工具会在你确认后关闭当前窗口，随后原地替换 exe。
- 如果当前是 Python 脚本运行模式，只会提示更新，不会覆盖源码。

---

## 三、在 Fork 里怎么配置（必读）

### 3.1 合并工具（解决 Excel 冲突）

当 Git 合并产生冲突且冲突文件是 `.xlsx` 时，用本工具来选「保留本地」或「保留线上」。

1. 打开 **Fork** → **Preferences（设置）** → **Integration** → **Merge Tool**。
2. **Merger** 选 **Custom**。
3. **Merger Path** 填：`ExcelMergeFork.exe` 的**完整路径**（例如：`D:\Tools\ForkExcelMergeTool\ExcelMergeFork.exe`）。
4. **Arguments** 填：`$LOCAL,$BASE,$REMOTE,$MERGED`
   （工具也兼容空格分隔：`$LOCAL $BASE $REMOTE $MERGED`。）

保存后，在 Fork 里对冲突的 Excel 执行「解决冲突」时会自动用本工具打开。

### 3.2 对比工具（查看 Excel 差异）

用于在 Fork 里对比「当前分支」和「其它分支」的同一 Excel 文件差异。

1. **Fork** → **Preferences** → **Integration** → **External Diff Tool**。
2. **Diff Tool** 选 **Custom**。
3. **Diff Tool Path** 填：同上，`ExcelMergeFork.exe` 的完整路径。
4. **Arguments** 填：`"$REMOTE" "$LOCAL"`。

配置好后，在 Fork 里对某 Excel 使用「对比」时会生成 `{文件名}_compare.xlsx` 并自动打开。

---

## 四、合并模式：一步步怎么操作

当你在 Fork 里对冲突的 Excel 使用「解决冲突」后，会弹出 **Excel 三向合并** 窗口，按下面步骤即可。

### 4.1 看明白「本地」和「线上」

- **本地 (Local)**：当前分支的版本（左边），一般是「你自己的修改」。
- **线上 (Remote)**：被合并进来的分支的版本（右边），一般是「别人或别的分支的修改」。

窗口里会显示两边的提交信息（Hash、提交人、时间、事件），方便判断哪边是哪次修改。

### 4.2 处理冲突列表

- 中间表格会列出**所有冲突行**（Sheet 名 + 关键列 Key + 当前选择）。
- 对每一行，你可以选：
  - **取本地**：保留当前分支这一行的内容；
  - **取线上**：保留被合并分支这一行的内容。
- 操作方式：
  - 在表格里**选中一行**，再点下方的「**取本地**」或「**取线上**」；
  - 或直接在表格里点选该行，然后点对应按钮即可。
- 若没有冲突，表格为空，可直接进行 4.3。

### 4.3 生成结果并确认

1. 选完所有冲突后，点击「**生成合并结果**」。
2. 用「**打开合并文件**」检查合并后的 Excel 是否正确；需要额外留档时可点「**手动保存备份**」。
3. 确认无误后，点击「**确认无误并解决冲突**」—— 工具会帮你在 Git 里标记冲突已解决（如 `git add` 等），Fork 会使用合并后的文件。
4. 若不想保留本次合并，点「**取消**」即可。

### 4.4 备份文件在哪

合并窗口可以在「备份根目录」里自定义保存位置；点「选择目录」后会保存到本地配置。合并完成后会自动备份，也可以点「手动保存备份」再保存一份。

备份目录结构为：

```text
备份根目录\项目名\时间戳\
```

若未设置备份根目录，默认使用 **MERGED 文件所在目录**下的 `MergeExcelBackup`。每次备份目录里包含：

- `{合并文件名}_local.xlsx`：本地版本
- `{合并文件名}_remote.xlsx`：线上版本  
- `{合并文件名}_merged.xlsx`：本次合并结果

「打开备份目录」会打开本次备份所在目录；若还没生成结果，则打开当前项目的备份目录。

---

## 五、对比模式

- 在 Fork 里对某 Excel 使用「对比」时，会调用本工具。
- 工具会生成 **`{原文件名}_compare.xlsx`**，放在原文件同目录，并尝试用系统默认程序打开。
- 对比表里会标出：仅 A 有、仅 B 有、修改、相同。颜色含义见 [颜色与场景说明.md](颜色与场景说明.md)。

---

## 六、不通过 Fork 使用（命令行）

适合本机测试或脚本调用。

**合并（4 个参数）：**

```text
ExcelMergeFork.exe <本地文件> <基准文件> <线上文件> <输出合并文件>
python Scripts\MergeExcelFork.py <本地> <基准> <线上> <输出>
```

例如：

```text
ExcelMergeFork.exe D:\repo\Test.xlsx D:\repo\Base.xlsx D:\repo\Remote.xlsx D:\repo\Merged.xlsx
```

**对比（2 个参数）：**

```text
ExcelMergeFork.exe <线上文件REMOTE> <本地文件LOCAL>
python Scripts\MergeExcelFork.py <线上文件REMOTE> <本地文件LOCAL>
```

会在本地文件同目录生成 `{本地文件名}_compare.xlsx` 并自动打开。

---

## 七、常见问题

### 7.1 界面显示「0 处冲突」但我确定有冲突

可能原因与处理：

1. **首列是“行关键列”**：工具用**每个 Sheet 的第一列**作为行的唯一标识（Key）。若第一列是空的或不是关键列，可能识别不到冲突。请保证三个文件（本地、基准、线上）的**第一列都是同一套关键值**（如编号、ID、姓名等）。
2. **用 exe 时请用最新版**：若你用的是打包的 `ExcelMergeFork.exe`，请用**重新打包后的最新 exe**（运行项目里的 `一键打包.bat` 生成），否则可能仍是旧逻辑。
3. **看日志**：与 exe 同目录下的 **`MergeExcelFork.log`** 里会有本次合并的统计，例如：  
   `[MERGE] Sheet=Sheet1 行数 local=10 base=10 remote=10 all_keys=10 冲突=2`  
   可根据「行数」「all_keys」「冲突」判断是否读到了数据、算出了冲突。

### 7.2 底部按钮看不到 / 要拖窗口才能看到

当前版本已把底部操作栏固定在最下方，正常窗口大小下应能直接看到「生成合并结果」「打开合并文件」「确认无误并解决冲突」「取消」。若仍被遮挡，可把窗口拉大一些或最大化。

### 7.3 两边明显不同，但冲突表里没有或很少

工具会把**所有差异**都列出来，分两类：

- **需选择**：本地和线上**都有**这一行且内容不同，需要你选「取本地」或「取线上」。
- **仅本地有 / 仅线上有**：这一行只在一边出现，合并时会自动保留有的一方，表中会标成「将保留本地」或「将保留线上」，无需再选。

所以若你看到的是「仅本地有 A、D」「仅线上有 …」而没有「需选择」的行，说明两边只是**谁多了哪几行**的差异，没有**同一行两边改得不一样**的冲突，表里会把这些“仅一方有”的行也列出来，方便你核对。

另外：**同一 key（首列相同）出现多行时，只按一行参与合并**（后一行覆盖前一行）。例如本地有 2 行 key 都是「D」、线上有 3 行 key 都是「B」，表里只会各显示一行 D、一行 B，合并结果里每种 key 也只会保留一行。若需要保留多行，请把首列做成唯一 key（如加行号）。

### 7.4 合并后少了一列 / 表头格式乱了

若本地、基准、线上三份表**列数不一致**（例如本地 7 列、线上 6 列），工具会按**三份表里最大的列数**统一读取和写出，缺列的一方用空单元格补齐，这样合并结果不会少列、表头也不会错位。请用最新版本或重新打包 exe 后再试。

### 7.5 哪些 Sheet 会参与合并 / 对比

- 默认会处理**所有 Sheet**（包括常见的 Sheet1、Sheet2）。
- 只有**以 `#` 开头的表名**会被跳过（可用于放说明、临时表等不参与合并的内容）。

### 7.6 闪退或报错

请查看与 exe（或脚本）同目录下的 **`MergeExcelFork.log`**，里面有错误堆栈，便于排查或反馈问题。

---

## 八、退出码与日志（给脚本/自动化用）

- **0**：成功  
- **1**：参数或文件错误  
- **2**：合并/对比过程异常  

日志文件：与 exe/脚本同目录的 **`MergeExcelFork.log`**。

---

## 九、给开发者：打包与测试

- **一键打包**：双击 **`一键打包.bat`**，会自动安装依赖、打包 exe、并做一次快速验证。
- **一键打包并发布 Release**：执行 `一键打包.bat --gui --test --release`。需要本机已安装 GitHub CLI，并先执行过 `gh auth login`。
- **仅发布当前 exe**：执行 `powershell -NoProfile -ExecutionPolicy Bypass -File Scripts\publish-release.ps1`。
- **手动打包**：`pip install pyinstaller` 后执行 `build_exe.bat`。
- **测试合并**：双击 `run_quick_test.bat`，或手动：  
  `run_merge.bat TestData\local.xlsx TestData\base.xlsx TestData\remote.xlsx TestData\_output\merged.xlsx`
- **测试对比**：`run_compare.bat TestData\local.xlsx TestData\remote.xlsx`

### 9.1 Release 发布规则

发布脚本会读取 `Scripts\version.py` 的 `__version__`，创建形如 `v2.54` 的 GitHub Release，并上传：

- `ExcelMergeFork.exe`
- `ExcelMergeFork.exe.sha256`

应用内更新检查固定读取仓库 `ZhouJun2303/ForkExcelMergeTool` 的 latest release，并查找资产名 `ExcelMergeFork.exe`。如果要改仓库或资产名，请同步修改 `Scripts\config.py`。

---

## 十、项目结构（开发者）

**规范：所有脚本放在 `Scripts` 文件夹下。** 根目录仅保留启动器、批处理、配置与文档。

| 路径 | 职责 |
|------|------|
| **MergeExcelFork.py**（根目录） | 启动器：将 Scripts 加入路径后调用 Scripts.MergeExcelFork.main() |
| **Scripts/MergeExcelFork.py** | 入口：解析参数，启动合并/对比 GUI 或回退命令行 |
| **Scripts/config.py** | 全局常量：日志文件名、默认备份目录名、对比后缀、Sheet 跳过前缀 |
| **Scripts/log_util.py** | 日志：写日志到文件、解析日志路径 |
| **Scripts/backup_util.py** | 合并备份：读取备份根目录配置，按项目/时间创建备份目录 |
| **Scripts/excel_io.py** | Excel 读写与行/Key 抽象：加载 Sheet、合并格取值、Key 规范化、行相等判断 |
| **Scripts/merge_core.py** | 三向合并核心：无 GUI 时的合并与备份（命令行回退用） |
| **Scripts/compare_core.py** | 二向对比核心：计算差异、生成对比 Excel |
| **Scripts/git_util.py** | Git 操作：冲突解决后 git add、清理临时文件、兼容清理旧版扁平备份、获取提交信息 |
| **Scripts/conflict.py** | 冲突检测：三份 Excel 比较得到冲突项与每 Sheet 数据（供合并 GUI） |
| **Scripts/gui_common.py** | GUI 公共：日志到状态栏、颜色图例、打开文件、统一样式 |
| **Scripts/merge_gui.py** | 合并窗口：冲突列表、生成合并结果（以本地为底）、确认并解决冲突 |
| **Scripts/diff_gui.py** | 对比窗口：差异列表、生成并打开对比 Excel |
| **Scripts/ExcelMergeGUI.py** | 向后兼容：对外仍暴露 MergeWindow/DiffWindow，直接运行可预览合并界面（需根目录 TestData） |

---

把本 README 和 `ExcelMergeFork.exe`（或整个文件夹）分发给同事时，大家按 **第三节（Fork 配置）** 和 **第四节（合并操作）** 即可完成配置与日常使用。
