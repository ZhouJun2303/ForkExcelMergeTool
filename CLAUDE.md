# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel three-way merge and two-way comparison tool, compatible with **Fork** client's Merge Tool and Diff Tool. Resolves Git merge conflicts for Excel files by allowing users to choose "Local" or "Remote" for each conflicting row/column.

## Quick Commands

**Run merge (4 args):**
```bash
python MergeExcelFork.py <local> <base> <remote> <merged>
python Scripts\MergeExcelFork.py <local> <base> <remote> <merged>
```

**Git merge driver entry (called by Git):**
```bash
ExcelMergeFork.exe --git-merge-driver <base> <current> <other> <repo-path>
```

**Run compare (2 args):**
```bash
python MergeExcelFork.py <fileA> <fileB>
```

**Build exe:**
```bash
package.bat      # builds full and lite exe
build_exe.bat    # bundled-Python exe only
```

**Run tests:**
```bash
run_quick_test.bat
run_merge_mode_tests.bat
```

## Architecture

### Entry Points
- `MergeExcelFork.py` (root) - Launcher using `runpy` to execute `Scripts/MergeExcelFork.py`
- `Scripts/MergeExcelFork.py` - Main entry: parses args, launches merge GUI or compare GUI, falls back to CLI

### Core Modules (Scripts/)
| Module | Responsibility |
|--------|----------------|
| `config.py` | Global constants: log file, default backup dir name, compare suffix, sheet skip prefix |
| `log_util.py` | Logging to file, lock management for single-instance GUI |
| `backup_util.py` | Merge backup helpers: persisted root dir, project/timestamp path layout, file copies |
| `excel_io.py` | Excel read/write abstraction: row loading, key normalization, merged cell handling, row equality |
| `conflict.py` | Conflict detection: compares LOCAL/BASE/REMOTE, returns conflicts and sheet data |
| `merge_core.py` | Three-way merge logic with 5 modes (A-E) and options-driven pipeline |
| `compare_core.py` | Two-way comparison: computes diffs, generates contrast Excel |
| `git_util.py` | Git operations: `git add`, cleanup temp files, compatibility cleanup for old flat backup files |
| `git_merge_driver.py` | Git merge driver adapter: copies Git inputs to an isolated temp dir, opens merge GUI, atomically writes confirmed result back to `%A` without `git add` |
| `merge_gui.py` | Merge window: conflict list, option checkboxes, generate merged result |
| `diff_gui.py` | Compare window: diff list, generate and open contrast Excel |
| `gui_common.py` | Shared GUI utilities: status bar logging, color legends, file opening |
| `ExcelMergeGUI.py` | Backward compatibility: exposes MergeWindow/DiffWindow for direct runs |

### Merge Modes
- **Mode A**: Insert new rows (base + new rows from other side, inserted at prefix group end)
- **Mode B**: Insert new columns (base + new columns from other side)
- **Mode C**: Insert new sheets (append sheets that exist only in other side)
- **Mode D**: Conflict resolution (apply user choices for conflicting rows/columns)
- **Mode E**: Smart merge (A → B → C → D pipeline)

### Options-driven Pipeline
Options (persisted to `merge_options.json`):
- **A**: Keep rows unchanged (skip inserting new rows)
- **B**: Keep columns unchanged (skip inserting new columns)
- **C**: Delete rows (remove rows that exist in base but not in other)
- **D**: Delete columns (remove columns that exist in base but not in other)
- **E**: Add new sheets
- **F**: Delete sheets
- **G**: Resolve conflicts
- **backup_root_dir**: Optional custom backup root. Backups are written as `root/project/timestamp__excel__commit-labels/*__{local|remote|merged}__commit-label.xlsx`; empty uses `MERGED` directory + `MergeExcelBackup`. Long file names are shortened with a stable hash.
- **Merge/diff formats**: full merge/diff parsing supports `.xlsx` and `.xltx`; macro formats such as `.xlsm`/`.xltm` should use quick backup mode to avoid losing VBA on save.

### Key Conventions
- **First column as Key**: Each sheet's first column is the unique row identifier
- **Key normalization**: Numbers like `1` and `1.0` are treated as the same key
- **Sheet filtering**: Sheets starting with `#` are skipped (e.g., `#说明`)
- **Color coding** (merge output):
  - Green (#CCFFCC): New rows (key not in BASE)
  - Yellow (#FFFF99): Modified rows (key in BASE, content changed)
  - Red (#FFCCCC): Conflicts (LOCAL ≠ REMOTE, both differ from BASE)

### Data Flow
1. **Merge**: `MergeExcelFork.py` → `merge_gui.py` (GUI) or `merge_core.py` (CLI) → `conflict.py` for detection → `excel_io.py` for read/write → `backup_util.py` for backups → `git_util.py` for cleanup
2. **Compare**: `MergeExcelFork.py` → `diff_gui.py` (GUI) or `compare_core.py` (CLI) → `excel_io.py` → writes `{filename}_compare.xlsx`
3. **Git driver**: `MergeExcelFork.py --git-merge-driver` → `git_merge_driver.py` → isolated temp workbooks → `merge_gui.py` → confirmed result replaces Git `%A`; no `git add` or Git temp cleanup.

### Exit Codes
- **0**: Success
- **1**: Parameter or file error
- **2**: Merge/compare exception

### Logging
Log file: `MergeExcelFork.log` in the same directory as exe/script. Contains merge statistics per sheet (row counts, conflict counts).
