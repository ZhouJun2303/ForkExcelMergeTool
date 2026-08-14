# Repository Guidelines

## Project Structure
Windows Excel three-way merge and two-way diff tool for Fork. Application code is C# / .NET 8 under `src/`:

- `ExcelMergeFork.Core` — workbook session, conflict, merge pipeline, compare, backup, Git/Fork
- `ExcelMergeFork.App` — WPF Fluent UI, published as `ExcelMergeFork.exe`
- `ExcelMergeFork.Tests` — xUnit against `TestData/` fixtures

Do not reintroduce the retired Python `Scripts/` / Tkinter / lite launcher stack.

## Commands
- `dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release`
- `package.bat` — test, publish self-contained exe to repo root
- `ExcelMergeFork.exe` — settings; 4 paths merge; 2 paths compare (remote then local)
- `install_fork_integration.bat` / `install_git_integration.bat`

## Style
C# nullable enabled, 4-space indent. Keep merge semantics (first-column key, `#` skip sheets, A–G options). User-facing strings may be Chinese.

## Testing
No pytest. Use `dotnet test` and existing `TestData/mode_*.xlsx`. Do not overwrite `merge_options.json` unless the task requires it.

## Commits
Conventional prefixes: `feat:`, `fix:`. PRs should list `dotnet test` and mention merge modes if behavior changes.
