# CLAUDE.md

Excel three-way merge and two-way comparison tool for Fork. C# / .NET 8 + WPF. Python/Tkinter sources have been removed.

## Commands

```bash
dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release
package.bat
ExcelMergeFork.exe <local> <base> <remote> <merged>
ExcelMergeFork.exe <remote> <local>
ExcelMergeFork.exe --git-merge-driver <base> <current> <other> <repo-path>
```

## Layout

| Project | Role |
|---------|------|
| `src/ExcelMergeFork.Core` | Key normalize, session, conflict, merge A–G, compare, backup, Git/Fork |
| `src/ExcelMergeFork.App` | Fluent settings / merge / diff / startup / backup windows |
| `src/ExcelMergeFork.Tests` | xUnit |
| `TestData/` | Mode A–E workbooks |

## Conventions

- First column is row key; `1` and `1.0` are the same key
- Sheets starting with `#` are skipped
- Merge/diff parse `.xlsx` / `.xltx`; macros use quick backup
- Git driver writes back `%A` only, no `git add`
- Exit codes: 0 success, 1 arg/file, 2 exception
- Log: `MergeExcelFork.log` next to the exe
