# Repository Guidelines

## Project Structure & Module Organization
This is a Windows-focused Python tool for Excel three-way merge and two-way diff, designed for Fork integration. The root `MergeExcelFork.py` is a thin launcher. Main code lives in `Scripts/`: `MergeExcelFork.py` handles arguments, `merge_core.py` and `compare_core.py` contain merge/diff logic, `merge_gui.py` and `diff_gui.py` provide Tkinter UI, and `excel_io.py` centralizes workbook access. The lightweight non-bundled launcher lives in `Tools/lite_launcher/`. Test fixtures and test generators are in `TestData/`. Build output appears in `build/`, `dist/`, and the root `ExcelMergeFork.exe` / `ExcelMergeFork-lite.exe`; avoid committing regenerated binaries or logs unless intentionally releasing them.

## Build, Test, and Development Commands
- `pip install -r requirements.txt` installs runtime dependencies (`openpyxl`).
- `python MergeExcelFork.py <local> <base> <remote> <merged>` runs merge mode through the root launcher.
- `python MergeExcelFork.py <fileA> <fileB>` runs compare mode and writes a `*_compare.xlsx` file.
- `run_quick_test.bat` generates sample data if needed and runs a quick merge into `TestData\_output`.
- `run_merge_mode_tests.bat` runs mode A-E merge checks.
- `package.bat --test` bumps the version, builds both `ExcelMergeFork.exe` and `ExcelMergeFork-lite.exe`, and runs merge mode tests.
- `build_exe.bat --test` builds only the bundled-Python PyInstaller exe and runs merge mode tests.

## Coding Style & Naming Conventions
Use Python 3.7+ compatible code, 4-space indentation, and standard library modules where practical. Keep filenames and module names lowercase with underscores, matching the existing `Scripts/*.py` style. Prefer explicit helper functions for workbook, path, and Git behavior instead of duplicating logic across GUI modules. User-facing strings may be Chinese, but code comments should be brief and clarify non-obvious merge behavior only.

## Testing Guidelines
There is no formal pytest suite. Validate changes with the batch scripts above, especially `run_merge_mode_tests.bat` for merge logic. When changing Excel parsing or merge behavior, regenerate fixtures with `python TestData\gen_merge_mode_tests.py` if fixture structure changes, then inspect files under `TestData\_output`.

## Commit & Pull Request Guidelines
Recent commits use short Conventional Commit prefixes, for example `feat:优化 diff` and `fix:merge bug`. Follow `feat:`, `fix:`, or another clear type, with a concise Chinese or English summary. Pull requests should describe the user-visible behavior, list verification commands, mention affected merge modes, and attach screenshots when GUI layout changes.

## Agent-Specific Notes
Do not overwrite local config such as `merge_options.json` or generated logs unless the task requires it. Preserve existing uncommitted work, and keep repository guidance aligned with `README.md` and `CLAUDE.md`.
