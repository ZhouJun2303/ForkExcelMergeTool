# ExcelMergeFork 源码

.NET 8 解决方案。仓库根目录的 `ExcelMergeFork.exe` 由 `package.bat` 发布，双击即新界面。

```bat
dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release
dotnet run --project src\ExcelMergeFork.App\ExcelMergeFork.App.csproj
package.bat
run_web.bat
```

`run_web.bat` 在浏览器打开 `http://127.0.0.1:5088`。界面只在仓库根目录 `web/`。`ExcelMergeFork.exe` 用 WebView2 加载同一份页面，打包时会把 `web/` 复制到 exe 旁边。
