/*
 * OneNote 导出工具 —— 图形界面启动器
 *
 * 这是一个极小的原生 Windows 程序，唯一职责是把 scripts\export-gui.ps1
 * 用系统自带的 Windows PowerShell 拉起来，并保证不出现黑色控制台窗口。
 *
 * 为什么要单独写个启动器，而不是直接双击 .cmd 或 .ps1：
 *   - .cmd / .bat 会被资源管理器的附件管理器拦下来弹「无法验证发布者」，
 *     而且默认会闪一个控制台窗口。资源管理器只对高风险扩展名做这道检查，
 *     启动 exe 本身不弹。
 *   - .ps1 不能直接双击（默认执行策略是 Restricted，双击只会静默失败）。
 *   - 顺带解决了图标问题：脚本文件没法自带图标，exe 可以。
 *
 * 用原生代码而不是 PowerShell + ps2exe 之类打包，是为了让产物不依赖任何
 * 运行时、体积极小，也不会被某些杀软当成「脚本打包器」误报。
 *
 * 编译：由 GitHub Actions 在推送 tag 时自动完成，见
 * .github/workflows/release.yml。仓库里不存放 exe 二进制，改这里只需推一个
 * tag，CI 会编译并把 zip 发到 Release。
 */

/* -municode 已经在命令行上定义了 UNICODE/_UNICODE，这里加守卫避免重定义警告 */
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <wchar.h>
#include <stdio.h>

/* MessageBoxW 在 user32 里。MinGW 默认就链接它，MSVC 不会 —— 必须显式要求，
   否则报 LNK2019: unresolved external symbol __imp_MessageBoxW。 */
#ifdef _MSC_VER
#pragma comment(lib, "user32.lib")
#endif

#define APP_TITLE L"OneNote 导出工具"

/* 弹一个错误框。启动阶段的失败没有别的途径能让用户看到原因 —— 这是个
   GUI 子系统的程序，没有控制台可以输出。 */
static void ShowError(const wchar_t *message, const wchar_t *detail)
{
    wchar_t buf[4096];
    if (detail && detail[0]) {
        _snwprintf(buf, 4096, L"%s\n\n%s", message, detail);
    } else {
        _snwprintf(buf, 4096, L"%s", message);
    }
    buf[4095] = L'\0';
    MessageBoxW(NULL, buf, APP_TITLE, MB_ICONERROR | MB_OK | MB_SETFOREGROUND);
}

/* 取本 exe 所在目录。用宽字符版本，路径里带中文也不会出问题。 */
static BOOL GetOwnDirectory(wchar_t *out, DWORD cap)
{
    DWORD len = GetModuleFileNameW(NULL, out, cap);
    if (len == 0 || len >= cap) return FALSE;      /* 0=失败，>=cap=被截断 */

    wchar_t *slash = wcsrchr(out, L'\\');
    if (!slash) return FALSE;
    *slash = L'\0';
    return TRUE;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                    PWSTR pCmdLine, int nCmdShow)
{
    (void)hInstance; (void)hPrevInstance; (void)pCmdLine; (void)nCmdShow;

    wchar_t baseDir[32768];
    if (!GetOwnDirectory(baseDir, 32768)) {
        ShowError(L"无法确定程序所在位置。", L"请把本程序放在解压后的文件夹里运行。");
        return 1;
    }

    /* 目标脚本固定在本 exe 同级目录下的 scripts\ 里 */
    wchar_t scriptPath[32768];
    _snwprintf(scriptPath, 32768, L"%s\\scripts\\export-gui.ps1", baseDir);
    scriptPath[32767] = L'\0';

    if (GetFileAttributesW(scriptPath) == INVALID_FILE_ATTRIBUTES) {
        wchar_t detail[32768 + 256];
        _snwprintf(detail, 32768 + 256,
                   L"找不到脚本文件：\n%s\n\n"
                   L"请确认本程序与 scripts 文件夹在同一目录下（即解压后的完整文件夹）。",
                   scriptPath);
        detail[32768 + 255] = L'\0';
        ShowError(L"文件不完整。", detail);
        return 1;
    }

    /* 找 Windows PowerShell 5.1 —— 系统自带，一定存在。
       找不到再退回 PATH 上的 pwsh.exe（PowerShell 7）。 */
    wchar_t winDir[MAX_PATH];
    wchar_t psExe[32768] = L"";

    if (GetWindowsDirectoryW(winDir, MAX_PATH)) {
        _snwprintf(psExe, 32768,
                   L"%s\\System32\\WindowsPowerShell\\v1.0\\powershell.exe", winDir);
        psExe[32767] = L'\0';
        if (GetFileAttributesW(psExe) == INVALID_FILE_ATTRIBUTES) psExe[0] = L'\0';
    }

    if (!psExe[0]) {
        wchar_t found[MAX_PATH];
        if (SearchPathW(NULL, L"pwsh.exe", NULL, MAX_PATH, found, NULL)) {
            wcscpy(psExe, found);
        } else {
            ShowError(L"找不到 PowerShell。",
                      L"本工具需要 Windows 自带的 Windows PowerShell 5.1，"
                      L"在 Windows 7 及以上版本中随系统提供。");
            return 1;
        }
    }

    /* -STA 是 WPF 界面必需的（OneNote COM 与剪贴板都要求单线程单元）。
       -ExecutionPolicy Bypass 只作用于本次调用，不改变系统设置。 */
    wchar_t cmdLine[32768 * 2 + 512];
    _snwprintf(cmdLine, 32768 * 2 + 512,
               L"\"%s\" -NoProfile -STA -ExecutionPolicy Bypass -File \"%s\"",
               psExe, scriptPath);
    cmdLine[32768 * 2 + 511] = L'\0';

    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    /* CREATE_NO_WINDOW：powershell.exe 是控制台程序，不加这个标志会先弹一个
       黑框再出现界面。界面本身由 WPF 自己创建，不受影响。 */
    if (!CreateProcessW(NULL, cmdLine, NULL, NULL, FALSE,
                        CREATE_NO_WINDOW, NULL, baseDir, &si, &pi)) {
        wchar_t detail[256];
        _snwprintf(detail, 256, L"启动 PowerShell 失败，错误码 %lu。", GetLastError());
        detail[255] = L'\0';
        ShowError(L"无法启动界面。", detail);
        return 1;
    }

    /* 等界面关闭再退出：这样本进程的生命周期与界面一致，行为可预期。
       本进程没有窗口，用户不会看到它。 */
    WaitForSingleObject(pi.hProcess, INFINITE);

    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    return (int)exitCode;
}
