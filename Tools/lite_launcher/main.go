package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	_ "embed"
	"encoding/hex"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"syscall"
	"unsafe"
)

//go:embed payload.zip
var payloadZip []byte

var appVersion = "dev"

const (
	appTitle              = "ExcelMergeFork-lite"
	updateAssetName       = "ExcelMergeFork-lite.exe"
	updateSha256AssetName = "ExcelMergeFork-lite.exe.sha256"
	createNoWindow        = 0x08000000
)

type pythonCandidate struct {
	name    string
	args    []string
	version string
}

func main() {
	selfTest := hasArg("--lite-self-test")

	exePath, err := os.Executable()
	if err != nil {
		fail(selfTest, "无法读取当前 exe 路径: "+err.Error())
	}
	exePath, _ = filepath.Abs(exePath)
	homeDir := filepath.Dir(exePath)

	runtimeDir, err := prepareRuntime()
	if err != nil {
		fail(selfTest, "轻量版运行文件准备失败: "+err.Error())
	}

	python, err := findPython()
	if err != nil {
		fail(selfTest, missingPythonMessage(err.Error()))
	}

	if err := requireImport(python, "openpyxl"); err != nil {
		fail(selfTest, missingDependencyMessage("openpyxl", "python -m pip install openpyxl"))
	}
	if err := requireImport(python, "tkinter"); err != nil {
		fail(selfTest, missingDependencyMessage("tkinter", "请安装 Windows 官方 Python，并确保安装 Tcl/Tk and IDLE 组件。"))
	}

	if selfTest {
		fmt.Printf("OK: %s runtime=%s python=%s %s\n", appTitle, runtimeDir, python.commandLine(), python.version)
		return
	}

	code, err := runApp(python, runtimeDir, homeDir, exePath, filteredArgs(os.Args[1:]))
	if err != nil {
		fail(false, "启动 Python 运行时失败: "+err.Error())
	}
	os.Exit(code)
}

func hasArg(flag string) bool {
	for _, arg := range os.Args[1:] {
		if arg == flag {
			return true
		}
	}
	return false
}

func filteredArgs(args []string) []string {
	result := make([]string, 0, len(args))
	for _, arg := range args {
		if arg != "--lite-self-test" {
			result = append(result, arg)
		}
	}
	return result
}

func prepareRuntime() (string, error) {
	if len(payloadZip) == 0 {
		return "", errors.New("payload.zip 为空，请重新打包")
	}
	sum := sha256.Sum256(payloadZip)
	hash := hex.EncodeToString(sum[:])
	shortHash := hash[:12]

	base := os.Getenv("LOCALAPPDATA")
	if strings.TrimSpace(base) == "" {
		cacheDir, err := os.UserCacheDir()
		if err != nil {
			base = os.TempDir()
		} else {
			base = cacheDir
		}
	}

	runtimeDir := filepath.Join(base, "ExcelMergeFork", "lite-runtime", "v"+safeName(appVersion)+"-"+shortHash)
	markerPath := filepath.Join(runtimeDir, ".payload.sha256")
	entryPath := filepath.Join(runtimeDir, "MergeExcelFork.py")
	if marker, err := os.ReadFile(markerPath); err == nil && strings.TrimSpace(string(marker)) == hash {
		if _, err := os.Stat(entryPath); err == nil {
			return runtimeDir, nil
		}
	}

	if err := os.RemoveAll(runtimeDir); err != nil {
		return "", err
	}
	if err := os.MkdirAll(runtimeDir, 0755); err != nil {
		return "", err
	}
	if err := extractPayload(runtimeDir); err != nil {
		_ = os.RemoveAll(runtimeDir)
		return "", err
	}
	if err := os.WriteFile(markerPath, []byte(hash), 0644); err != nil {
		return "", err
	}
	return runtimeDir, nil
}

func safeName(text string) string {
	var b strings.Builder
	for _, r := range text {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == '.' || r == '-' || r == '_' {
			b.WriteRune(r)
		} else {
			b.WriteByte('_')
		}
	}
	if b.Len() == 0 {
		return "dev"
	}
	return b.String()
}

func extractPayload(runtimeDir string) error {
	reader, err := zip.NewReader(bytes.NewReader(payloadZip), int64(len(payloadZip)))
	if err != nil {
		return err
	}
	for _, file := range reader.File {
		target := filepath.Join(runtimeDir, file.Name)
		if !isSafeChild(runtimeDir, target) {
			return fmt.Errorf("payload 包含非法路径: %s", file.Name)
		}
		if file.FileInfo().IsDir() {
			if err := os.MkdirAll(target, 0755); err != nil {
				return err
			}
			continue
		}
		if err := os.MkdirAll(filepath.Dir(target), 0755); err != nil {
			return err
		}
		src, err := file.Open()
		if err != nil {
			return err
		}
		dst, err := os.OpenFile(target, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, file.Mode())
		if err != nil {
			_ = src.Close()
			return err
		}
		_, copyErr := io.Copy(dst, src)
		closeErr := dst.Close()
		_ = src.Close()
		if copyErr != nil {
			return copyErr
		}
		if closeErr != nil {
			return closeErr
		}
	}
	return nil
}

func isSafeChild(root, target string) bool {
	rootAbs, err := filepath.Abs(root)
	if err != nil {
		return false
	}
	targetAbs, err := filepath.Abs(target)
	if err != nil {
		return false
	}
	rel, err := filepath.Rel(rootAbs, targetAbs)
	if err != nil {
		return false
	}
	return rel != "." && !strings.HasPrefix(rel, ".."+string(filepath.Separator)) && rel != ".."
}

func findPython() (pythonCandidate, error) {
	candidates := []pythonCandidate{
		{name: "py", args: []string{"-3"}},
		{name: "python"},
		{name: "python3"},
	}
	var problems []string
	for _, candidate := range candidates {
		version, err := pythonVersion(candidate)
		if err != nil {
			problems = append(problems, candidate.commandLine()+": "+err.Error())
			continue
		}
		if !versionAtLeast(version, 3, 7) {
			problems = append(problems, candidate.commandLine()+": Python "+version+" 低于 3.7")
			continue
		}
		candidate.version = version
		return candidate, nil
	}
	if len(problems) == 0 {
		return pythonCandidate{}, errors.New("未找到 py/python/python3")
	}
	return pythonCandidate{}, errors.New(strings.Join(problems, "\n"))
}

func pythonVersion(candidate pythonCandidate) (string, error) {
	out, err := runPythonCapture(candidate, "import sys; print('%d.%d.%d' % sys.version_info[:3])")
	if err != nil {
		return "", err
	}
	version := strings.TrimSpace(out)
	if version == "" {
		return "", errors.New("无法读取 Python 版本")
	}
	return version, nil
}

func versionAtLeast(version string, major, minor int) bool {
	parts := strings.Split(version, ".")
	if len(parts) < 2 {
		return false
	}
	gotMajor, err := strconv.Atoi(parts[0])
	if err != nil {
		return false
	}
	gotMinor, err := strconv.Atoi(parts[1])
	if err != nil {
		return false
	}
	if gotMajor != major {
		return gotMajor > major
	}
	return gotMinor >= minor
}

func requireImport(candidate pythonCandidate, module string) error {
	code := fmt.Sprintf("import %s", module)
	_, err := runPythonCapture(candidate, code)
	return err
}

func runPythonCapture(candidate pythonCandidate, code string) (string, error) {
	args := append([]string{}, candidate.args...)
	args = append(args, "-c", code)
	cmd := exec.Command(candidate.name, args...)
	hideConsole(cmd)
	out, err := cmd.CombinedOutput()
	if err != nil {
		text := strings.TrimSpace(string(out))
		if text != "" {
			return "", errors.New(text)
		}
		return "", err
	}
	return string(out), nil
}

func runApp(candidate pythonCandidate, runtimeDir, homeDir, launcherExe string, passthrough []string) (int, error) {
	entry := filepath.Join(runtimeDir, "MergeExcelFork.py")
	args := append([]string{}, candidate.args...)
	args = append(args, entry)
	args = append(args, passthrough...)

	cmd := exec.Command(candidate.name, args...)
	cmd.Dir = homeDir
	cmd.Stdin = os.Stdin
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	cmd.Env = append(os.Environ(),
		"EXCEL_MERGE_FORK_HOME="+homeDir,
		"EXCEL_MERGE_FORK_RESOURCE_ROOT="+runtimeDir,
		"EXCEL_MERGE_FORK_LAUNCHER_EXE="+launcherExe,
		"EXCEL_MERGE_FORK_UPDATE_ASSET_NAME="+updateAssetName,
		"EXCEL_MERGE_FORK_UPDATE_SHA256_ASSET_NAME="+updateSha256AssetName,
	)
	hideConsole(cmd)

	err := cmd.Run()
	if err == nil {
		return 0, nil
	}
	var exitErr *exec.ExitError
	if errors.As(err, &exitErr) {
		return exitErr.ExitCode(), nil
	}
	return 2, err
}

func hideConsole(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{
		CreationFlags: createNoWindow,
	}
}

func (candidate pythonCandidate) commandLine() string {
	parts := append([]string{candidate.name}, candidate.args...)
	return strings.Join(parts, " ")
}

func missingPythonMessage(detail string) string {
	return "ExcelMergeFork-lite.exe 是不内置 Python 的轻量版，需要本机已有 Python 3.7+。\n\n" +
		"请先安装 Python 3.7 或更高版本，并勾选 Add Python to PATH。\n" +
		"安装后再执行：\n" +
		"python -m pip install openpyxl\n\n" +
		"当前检测结果：\n" + detail
}

func missingDependencyMessage(module, action string) string {
	return "ExcelMergeFork-lite.exe 检测到当前 Python 环境缺少依赖：" + module + "\n\n" +
		"请处理后重新运行：\n" + action
}

func fail(selfTest bool, message string) {
	if selfTest {
		fmt.Fprintln(os.Stderr, "ERROR: "+message)
		os.Exit(1)
	}
	fmt.Fprintln(os.Stderr, "ERROR: "+message)
	messageBox(appTitle, message)
	os.Exit(1)
}

func messageBox(title, text string) {
	user32 := syscall.NewLazyDLL("user32.dll")
	proc := user32.NewProc("MessageBoxW")
	proc.Call(
		0,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(text))),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(title))),
		uintptr(0x00000010),
	)
}
