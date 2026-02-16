#include <windows.h>
#include <shellapi.h>
#include <commdlg.h>
#include <strsafe.h>

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <cwctype>
#include <filesystem>
#include <cmath>
#include <limits>
#include <string>
#include <string_view>

namespace {

constexpr UINT kTrayIconId = 1;
constexpr UINT kTrayMessage = WM_APP + 1;
constexpr UINT_PTR kMonitorTimerId = 1;
constexpr UINT_PTR kBlinkTimerId = 2;
constexpr UINT kDefaultMonitorIntervalMs = 1500;
constexpr UINT kBlinkIntervalMs = 500;
constexpr UINT kMinMonitorIntervalMs = 500;
constexpr double kMinMonitorIntervalSeconds = static_cast<double>(kMinMonitorIntervalMs) / 1000.0;
constexpr double kMaxTimerSupportedSeconds = static_cast<double>((std::numeric_limits<UINT>::max)()) / 1000.0;
constexpr wchar_t kSingleInstanceMutexName[] = L"Local\\BackrestTrayWatcher.Singleton";

constexpr UINT kMenuSetLogPath = 1001;
constexpr UINT kMenuOpenLogFolder = 1002;
constexpr UINT kMenuOpenLogFile = 1003;
constexpr UINT kMenuAcknowledgeAlert = 1004;
constexpr UINT kMenuExit = 1005;
constexpr UINT kMenuSetMonitorInterval = 1101;
constexpr UINT_PTR kAcknowledgePopupTimerId = 3;
constexpr UINT kDefaultAcknowledgePopupDurationMs = 2500;
constexpr UINT kMinAcknowledgePopupDurationMs = 500;
constexpr UINT kMaxAcknowledgePopupDurationMs = 30000;
constexpr std::string_view kAlertKeyword = "\"logger\":";
constexpr int kAcknowledgePopupWidth = 420;
constexpr int kAcknowledgePopupHeight = 72;
constexpr int kAcknowledgePopupOffsetPx = 8;
constexpr wchar_t kAcknowledgePopupWindowClassName[] = L"BackrestWatcherAcknowledgePopupWindowClass";

struct AppState {
  HWND hwnd = nullptr;
  std::wstring configPath;
  std::wstring logPath;
  ULONGLONG acknowledgedOffset = 0;
  ULONGLONG lastOffset = 0;
  UINT monitorIntervalMs = kDefaultMonitorIntervalMs;
  bool hasAlert = false;
  bool blinkShowAlertIcon = true;
  UINT acknowledgePopupDurationMs = kDefaultAcknowledgePopupDurationMs;
  NOTIFYICONDATAW trayIcon = {};
  HICON normalIcon = nullptr;
  HICON alertIcon = nullptr;
  bool ownsNormalIcon = false;
  bool trayIconAdded = false;
  HWND acknowledgePopupHwnd = nullptr;
  HANDLE singleInstanceMutex = nullptr;
  UINT taskbarCreatedMessage = 0;
};

AppState g_state;

struct IntervalInputDialogState {
  UINT initialValueMs = 0;
  UINT resultValueMs = 0;
  bool accepted = false;
  HWND editControl = nullptr;
  HWND secondsCheckbox = nullptr;
  HWND minutesCheckbox = nullptr;
};

std::wstring ExeDirectory() {
  wchar_t path[MAX_PATH] = {};
  GetModuleFileNameW(nullptr, path, MAX_PATH);
  std::filesystem::path exePath(path);
  return exePath.parent_path().wstring();
}

std::wstring DefaultLogPath() {
  return ExeDirectory() + L"\\backrest.log";
}

std::wstring ConfigFilePath() {
  return ExeDirectory() + L"\\backrest_tray_watcher.ini";
}

std::wstring NormalIconPath() {
  return ExeDirectory() + L"\\BackrestTrayWatcher.ico";
}

bool ContainsAlertKeyword(std::string_view text) {
  return text.find(kAlertKeyword) != std::string_view::npos;
}

bool ScanFileRangeForAlerts(HANDLE file, ULONGLONG beginOffset, ULONGLONG endOffset) {
  if (endOffset <= beginOffset) {
    return false;
  }

  LARGE_INTEGER filePointer = {};
  filePointer.QuadPart = static_cast<LONGLONG>(beginOffset);
  if (!SetFilePointerEx(file, filePointer, nullptr, FILE_BEGIN)) {
    return false;
  }

  constexpr DWORD kBufferSize = 64 * 1024;
  constexpr size_t kOverlapSize = kAlertKeyword.size() > 0 ? (kAlertKeyword.size() - 1) : 0;
  char buffer[kBufferSize];
  std::string overlap;
  overlap.reserve(kOverlapSize);

  ULONGLONG remaining = endOffset - beginOffset;
  while (remaining > 0) {
    const DWORD toRead = static_cast<DWORD>(
        std::min<ULONGLONG>(remaining, static_cast<ULONGLONG>(kBufferSize)));
    DWORD bytesRead = 0;
    if (!ReadFile(file, buffer, toRead, &bytesRead, nullptr)) {
      return false;
    }
    if (bytesRead == 0) {
      break;
    }

    std::string chunk;
    chunk.reserve(overlap.size() + bytesRead);
    chunk.append(overlap);
    chunk.append(buffer, bytesRead);
    if (ContainsAlertKeyword(chunk)) {
      return true;
    }

    if (chunk.size() > kOverlapSize) {
      overlap.assign(chunk.end() - static_cast<std::string::difference_type>(kOverlapSize), chunk.end());
    } else {
      overlap = chunk;
    }

    remaining -= bytesRead;
  }

  return false;
}

void SaveLogPathToConfig(const std::wstring& logPath) {
  WritePrivateProfileStringW(L"watcher", L"log_path", logPath.c_str(), g_state.configPath.c_str());
}

UINT ClampMonitorInterval(UINT intervalMs) {
  if (intervalMs < kMinMonitorIntervalMs) {
    return kMinMonitorIntervalMs;
  }
  return intervalMs;
}

UINT ClampAcknowledgePopupDuration(UINT durationMs) {
  if (durationMs < kMinAcknowledgePopupDurationMs) {
    return kMinAcknowledgePopupDurationMs;
  }
  if (durationMs > kMaxAcknowledgePopupDurationMs) {
    return kMaxAcknowledgePopupDurationMs;
  }
  return durationMs;
}

void SaveMonitorIntervalToConfig(UINT intervalMs) {
  wchar_t intervalBuffer[32] = {};
  StringCchPrintfW(intervalBuffer, ARRAYSIZE(intervalBuffer), L"%u", intervalMs);
  WritePrivateProfileStringW(L"watcher", L"monitor_interval_ms", intervalBuffer, g_state.configPath.c_str());
}

void LoadMonitorIntervalFromConfig() {
  const UINT configuredInterval = GetPrivateProfileIntW(
      L"watcher",
      L"monitor_interval_ms",
      kDefaultMonitorIntervalMs,
      g_state.configPath.c_str());
  g_state.monitorIntervalMs = ClampMonitorInterval(configuredInterval);
}

void SaveAcknowledgePopupDurationToConfig(UINT durationMs) {
  const UINT clampedDurationMs = ClampAcknowledgePopupDuration(durationMs);
  wchar_t durationBuffer[32] = {};
  StringCchPrintfW(
      durationBuffer,
      ARRAYSIZE(durationBuffer),
      L"%.3f",
      static_cast<double>(clampedDurationMs) / 1000.0);
  WritePrivateProfileStringW(L"watcher", L"ack_popup_seconds", durationBuffer, g_state.configPath.c_str());
}

void LoadAcknowledgePopupDurationFromConfig() {
  wchar_t durationBuffer[64] = {};
  const DWORD charsRead = GetPrivateProfileStringW(
      L"watcher",
      L"ack_popup_seconds",
      L"",
      durationBuffer,
      static_cast<DWORD>(ARRAYSIZE(durationBuffer)),
      g_state.configPath.c_str());

  if (charsRead == 0) {
    g_state.acknowledgePopupDurationMs = kDefaultAcknowledgePopupDurationMs;
    SaveAcknowledgePopupDurationToConfig(g_state.acknowledgePopupDurationMs);
    return;
  }

  wchar_t* parseEnd = nullptr;
  const double durationSeconds = wcstod(durationBuffer, &parseEnd);
  while (parseEnd != nullptr && *parseEnd != L'\0' && iswspace(*parseEnd) != 0) {
    ++parseEnd;
  }
  if (!std::isfinite(durationSeconds) ||
      durationSeconds < 0.0 ||
      parseEnd == durationBuffer ||
      (parseEnd != nullptr && *parseEnd != L'\0')) {
    g_state.acknowledgePopupDurationMs = kDefaultAcknowledgePopupDurationMs;
    SaveAcknowledgePopupDurationToConfig(g_state.acknowledgePopupDurationMs);
    return;
  }

  const double minSeconds = static_cast<double>(kMinAcknowledgePopupDurationMs) / 1000.0;
  const double maxSeconds = static_cast<double>(kMaxAcknowledgePopupDurationMs) / 1000.0;
  const double clampedSeconds = (std::max)(minSeconds, (std::min)(durationSeconds, maxSeconds));
  g_state.acknowledgePopupDurationMs = ClampAcknowledgePopupDuration(
      static_cast<UINT>(std::llround(clampedSeconds * 1000.0)));

  if (std::fabs(clampedSeconds - durationSeconds) > 0.0005) {
    SaveAcknowledgePopupDurationToConfig(g_state.acknowledgePopupDurationMs);
  }
}

void SaveAcknowledgedOffsetToConfig(ULONGLONG offset) {
  wchar_t offsetBuffer[32] = {};
  StringCchPrintfW(offsetBuffer, ARRAYSIZE(offsetBuffer), L"%llu", offset);
  WritePrivateProfileStringW(L"watcher", L"ack_offset", offsetBuffer, g_state.configPath.c_str());
}

void LoadAcknowledgedOffsetFromConfig() {
  wchar_t offsetBuffer[32] = {};
  GetPrivateProfileStringW(
      L"watcher",
      L"ack_offset",
      L"0",
      offsetBuffer,
      static_cast<DWORD>(ARRAYSIZE(offsetBuffer)),
      g_state.configPath.c_str());

  wchar_t* parseEnd = nullptr;
  g_state.acknowledgedOffset = _wcstoui64(offsetBuffer, &parseEnd, 10);
  if (parseEnd == offsetBuffer) {
    g_state.acknowledgedOffset = 0;
  }
}

void LoadLogPathFromConfig() {
  wchar_t logPathBuffer[4096] = {};
  const DWORD charsRead = GetPrivateProfileStringW(
      L"watcher",
      L"log_path",
      L"",
      logPathBuffer,
      static_cast<DWORD>(ARRAYSIZE(logPathBuffer)),
      g_state.configPath.c_str());

  if (charsRead > 0) {
    g_state.logPath = logPathBuffer;
  } else {
    g_state.logPath = DefaultLogPath();
  }

  LoadMonitorIntervalFromConfig();
  LoadAcknowledgePopupDurationFromConfig();
  LoadAcknowledgedOffsetFromConfig();
}

bool AcquireSingleInstanceLock(bool* alreadyRunning) {
  if (alreadyRunning) {
    *alreadyRunning = false;
  }

  g_state.singleInstanceMutex = CreateMutexW(nullptr, FALSE, kSingleInstanceMutexName);
  if (!g_state.singleInstanceMutex) {
    return false;
  }

  if (GetLastError() == ERROR_ALREADY_EXISTS && alreadyRunning) {
    *alreadyRunning = true;
  }

  return true;
}

void ReleaseSingleInstanceLock() {
  if (g_state.singleInstanceMutex) {
    CloseHandle(g_state.singleInstanceMutex);
    g_state.singleInstanceMutex = nullptr;
  }
}

void RefreshTrayIconVisualState() {
  std::wstring tip = L"Backrest Watcher: ";
  tip += g_state.hasAlert ? L"WARNING/ERROR" : L"OK";

  StringCchCopyW(g_state.trayIcon.szTip, ARRAYSIZE(g_state.trayIcon.szTip), tip.c_str());
  g_state.trayIcon.hIcon = (g_state.hasAlert && g_state.blinkShowAlertIcon) ? g_state.alertIcon : g_state.normalIcon;
}

bool AddTrayIcon() {
  RefreshTrayIconVisualState();
  g_state.trayIcon.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
  if (!Shell_NotifyIconW(NIM_ADD, &g_state.trayIcon)) {
    return false;
  }
  g_state.trayIconAdded = true;
  g_state.trayIcon.uVersion = NOTIFYICON_VERSION_4;
  Shell_NotifyIconW(NIM_SETVERSION, &g_state.trayIcon);
  return true;
}

void UpdateTrayIcon() {
  RefreshTrayIconVisualState();
  g_state.trayIcon.uFlags = NIF_ICON | NIF_TIP;
  if (!Shell_NotifyIconW(NIM_MODIFY, &g_state.trayIcon)) {
    AddTrayIcon();
  }
}

LRESULT CALLBACK AcknowledgePopupWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  switch (message) {
    case WM_CREATE:
      SetTimer(hwnd, kAcknowledgePopupTimerId, g_state.acknowledgePopupDurationMs, nullptr);
      return 0;

    case WM_TIMER:
      if (wParam == kAcknowledgePopupTimerId) {
        DestroyWindow(hwnd);
      }
      return 0;

    case WM_LBUTTONDOWN:
    case WM_RBUTTONDOWN:
      DestroyWindow(hwnd);
      return 0;

    case WM_PAINT: {
      PAINTSTRUCT paintStruct = {};
      HDC dc = BeginPaint(hwnd, &paintStruct);
      RECT clientRect = {};
      GetClientRect(hwnd, &clientRect);

      HBRUSH backgroundBrush = CreateSolidBrush(RGB(255, 249, 230));
      FillRect(dc, &clientRect, backgroundBrush);
      DeleteObject(backgroundBrush);

      HPEN borderPen = CreatePen(PS_SOLID, 1, RGB(208, 186, 132));
      HGDIOBJ oldPen = SelectObject(dc, borderPen);
      HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(HOLLOW_BRUSH));
      Rectangle(dc, clientRect.left, clientRect.top, clientRect.right, clientRect.bottom);
      SelectObject(dc, oldBrush);
      SelectObject(dc, oldPen);
      DeleteObject(borderPen);

      RECT textRect = clientRect;
      textRect.left += 10;
      textRect.top += 8;
      textRect.right -= 10;
      textRect.bottom -= 8;
      SetBkMode(dc, TRANSPARENT);
      SetTextColor(dc, RGB(30, 30, 30));
      DrawTextW(
          dc,
          L"Backrest Watcher\r\nAcknowledged. Monitoring continues from current log position.",
          -1,
          &textRect,
          DT_LEFT | DT_TOP | DT_WORDBREAK | DT_NOPREFIX);

      EndPaint(hwnd, &paintStruct);
      return 0;
    }

    case WM_DESTROY:
      KillTimer(hwnd, kAcknowledgePopupTimerId);
      if (g_state.acknowledgePopupHwnd == hwnd) {
        g_state.acknowledgePopupHwnd = nullptr;
      }
      return 0;

    default:
      return DefWindowProcW(hwnd, message, wParam, lParam);
  }
}

bool EnsureAcknowledgePopupWindowClassRegistered() {
  static bool registered = false;
  if (registered) {
    return true;
  }

  WNDCLASSW popupClass = {};
  popupClass.lpfnWndProc = AcknowledgePopupWindowProc;
  popupClass.hInstance = GetModuleHandleW(nullptr);
  popupClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  popupClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_INFOBK + 1);
  popupClass.lpszClassName = kAcknowledgePopupWindowClassName;

  if (!RegisterClassW(&popupClass)) {
    if (GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
      return false;
    }
  }

  registered = true;
  return true;
}

void ShowAcknowledgeNotification() {
  if (!EnsureAcknowledgePopupWindowClassRegistered()) {
    return;
  }

  if (g_state.acknowledgePopupHwnd && IsWindow(g_state.acknowledgePopupHwnd)) {
    DestroyWindow(g_state.acknowledgePopupHwnd);
    g_state.acknowledgePopupHwnd = nullptr;
  }

  POINT cursorPos = {};
  if (!GetCursorPos(&cursorPos)) {
    return;
  }

  RECT workArea = {0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)};
  MONITORINFO monitorInfo = {};
  monitorInfo.cbSize = sizeof(monitorInfo);
  HMONITOR monitor = MonitorFromPoint(cursorPos, MONITOR_DEFAULTTONEAREST);
  if (monitor && GetMonitorInfoW(monitor, &monitorInfo)) {
    workArea = monitorInfo.rcWork;
  }

  int popupX = cursorPos.x - (kAcknowledgePopupWidth / 2);
  int popupY = cursorPos.y + kAcknowledgePopupOffsetPx;

  if (popupY + kAcknowledgePopupHeight > workArea.bottom) {
    popupY = cursorPos.y - kAcknowledgePopupHeight - kAcknowledgePopupOffsetPx;
  }

  popupX = std::max<int>(workArea.left, std::min<int>(popupX, workArea.right - kAcknowledgePopupWidth));
  popupY = std::max<int>(workArea.top, std::min<int>(popupY, workArea.bottom - kAcknowledgePopupHeight));

  HWND popupHwnd = CreateWindowExW(
      WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
      kAcknowledgePopupWindowClassName,
      L"",
      WS_POPUP,
      popupX,
      popupY,
      kAcknowledgePopupWidth,
      kAcknowledgePopupHeight,
      g_state.hwnd,
      nullptr,
      GetModuleHandleW(nullptr),
      nullptr);
  if (!popupHwnd) {
    return;
  }

  g_state.acknowledgePopupHwnd = popupHwnd;
  ShowWindow(popupHwnd, SW_SHOWNOACTIVATE);
  UpdateWindow(popupHwnd);
}

void ResetWatcherAndRescan() {
  g_state.lastOffset = 0;
  g_state.hasAlert = false;
  g_state.blinkShowAlertIcon = true;

  HANDLE file = CreateFileW(
      g_state.logPath.c_str(),
      GENERIC_READ,
      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
      nullptr,
      OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL,
      nullptr);

  if (file != INVALID_HANDLE_VALUE) {
    LARGE_INTEGER fileSize = {};
    if (GetFileSizeEx(file, &fileSize) && fileSize.QuadPart > 0) {
      const ULONGLONG currentSize = static_cast<ULONGLONG>(fileSize.QuadPart);
      if (g_state.acknowledgedOffset > currentSize) {
        g_state.acknowledgedOffset = 0;
        SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
      }
      g_state.hasAlert = ScanFileRangeForAlerts(file, g_state.acknowledgedOffset, currentSize);
      g_state.lastOffset = currentSize;
    }
    CloseHandle(file);
  }

  UpdateTrayIcon();
}

void MonitorLogFileOnce() {
  HANDLE file = CreateFileW(
      g_state.logPath.c_str(),
      GENERIC_READ,
      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
      nullptr,
      OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL,
      nullptr);

  if (file == INVALID_HANDLE_VALUE) {
    if (g_state.hasAlert) {
      g_state.hasAlert = false;
      g_state.blinkShowAlertIcon = true;
      UpdateTrayIcon();
    }
    g_state.lastOffset = 0;
    return;
  }

  LARGE_INTEGER fileSize = {};
  if (!GetFileSizeEx(file, &fileSize)) {
    CloseHandle(file);
    return;
  }

  const ULONGLONG newSize = static_cast<ULONGLONG>(fileSize.QuadPart);
  bool needIconRefresh = false;

  if (newSize < g_state.lastOffset) {
    if (g_state.acknowledgedOffset > newSize) {
      g_state.acknowledgedOffset = 0;
      SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
    }
    g_state.lastOffset = 0;
    g_state.hasAlert = false;
    g_state.blinkShowAlertIcon = true;
    needIconRefresh = true;
  }

  if (newSize > g_state.lastOffset) {
    if (ScanFileRangeForAlerts(file, g_state.lastOffset, newSize)) {
      if (!g_state.hasAlert) {
        g_state.hasAlert = true;
        g_state.blinkShowAlertIcon = true;
        needIconRefresh = true;
      }
    }
    g_state.lastOffset = newSize;
  }

  if (needIconRefresh) {
    UpdateTrayIcon();
  }

  CloseHandle(file);
}

void OpenLogFolder() {
  std::filesystem::path logPath(g_state.logPath);
  if (logPath.empty()) {
    return;
  }

  if (std::filesystem::exists(logPath)) {
    std::wstring args = L"/select,\"" + logPath.wstring() + L"\"";
    ShellExecuteW(nullptr, L"open", L"explorer.exe", args.c_str(), nullptr, SW_SHOWNORMAL);
    return;
  }

  std::filesystem::path folder = logPath.parent_path();
  if (folder.empty()) {
    folder = ExeDirectory();
  }
  ShellExecuteW(nullptr, L"open", folder.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

void OpenLogFile() {
  const HINSTANCE result = ShellExecuteW(nullptr, L"open", g_state.logPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(result) <= 32) {
    MessageBoxW(g_state.hwnd, L"Cannot open log file.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
  }
}

void AcknowledgeAlert() {
  g_state.acknowledgedOffset = g_state.lastOffset;
  SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
  g_state.hasAlert = false;
  g_state.blinkShowAlertIcon = true;
  UpdateTrayIcon();
  ShowAcknowledgeNotification();
}

void ApplyMonitorInterval() {
  if (!g_state.hwnd) {
    return;
  }
  KillTimer(g_state.hwnd, kMonitorTimerId);
  if (SetTimer(g_state.hwnd, kMonitorTimerId, g_state.monitorIntervalMs, nullptr) == 0) {
    MessageBoxW(g_state.hwnd, L"Cannot update log monitoring timer.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
  }
}

void SetMonitorInterval(UINT intervalMs) {
  g_state.monitorIntervalMs = ClampMonitorInterval(intervalMs);
  SaveMonitorIntervalToConfig(g_state.monitorIntervalMs);
  ApplyMonitorInterval();
}

LRESULT CALLBACK IntervalInputWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  auto* state = reinterpret_cast<IntervalInputDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

  switch (message) {
    case WM_CREATE: {
      auto* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
      state = reinterpret_cast<IntervalInputDialogState*>(createStruct->lpCreateParams);
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));

      CreateWindowExW(
          0, L"STATIC", L"Enter check interval value:",
          WS_CHILD | WS_VISIBLE,
          12, 12, 266, 20,
          hwnd, nullptr, nullptr, nullptr);

      state->editControl = CreateWindowExW(
          WS_EX_CLIENTEDGE, L"EDIT", L"",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
          12, 36, 266, 24,
          hwnd, reinterpret_cast<HMENU>(100), nullptr, nullptr);

      state->secondsCheckbox = CreateWindowExW(
          0, L"BUTTON", L"Seconds",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
          12, 66, 90, 20,
          hwnd, reinterpret_cast<HMENU>(101), nullptr, nullptr);

      state->minutesCheckbox = CreateWindowExW(
          0, L"BUTTON", L"Minutes",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
          110, 66, 90, 20,
          hwnd, reinterpret_cast<HMENU>(102), nullptr, nullptr);

      const bool useMinutes = (state->initialValueMs >= 60000 && (state->initialValueMs % 60000) == 0);
      SendMessageW(state->secondsCheckbox, BM_SETCHECK, useMinutes ? BST_UNCHECKED : BST_CHECKED, 0);
      SendMessageW(state->minutesCheckbox, BM_SETCHECK, useMinutes ? BST_CHECKED : BST_UNCHECKED, 0);

      wchar_t initialBuffer[32] = {};
      if (useMinutes) {
        StringCchPrintfW(initialBuffer, ARRAYSIZE(initialBuffer), L"%.3f", static_cast<double>(state->initialValueMs) / 60000.0);
      } else {
        StringCchPrintfW(initialBuffer, ARRAYSIZE(initialBuffer), L"%.3f", static_cast<double>(state->initialValueMs) / 1000.0);
      }
      SetWindowTextW(state->editControl, initialBuffer);

      CreateWindowExW(
          0, L"BUTTON", L"OK",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
          122, 96, 74, 26,
          hwnd, reinterpret_cast<HMENU>(IDOK), nullptr, nullptr);

      CreateWindowExW(
          0, L"BUTTON", L"Cancel",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
          204, 96, 74, 26,
          hwnd, reinterpret_cast<HMENU>(IDCANCEL), nullptr, nullptr);

      SendMessageW(state->editControl, EM_SETSEL, 0, -1);
      SetFocus(state->editControl);
      return 0;
    }

    case WM_COMMAND:
      if (LOWORD(wParam) == 101) {
        SendMessageW(state->secondsCheckbox, BM_SETCHECK, BST_CHECKED, 0);
        SendMessageW(state->minutesCheckbox, BM_SETCHECK, BST_UNCHECKED, 0);
        return 0;
      }
      if (LOWORD(wParam) == 102) {
        SendMessageW(state->secondsCheckbox, BM_SETCHECK, BST_UNCHECKED, 0);
        SendMessageW(state->minutesCheckbox, BM_SETCHECK, BST_CHECKED, 0);
        return 0;
      }
      if (LOWORD(wParam) == IDOK) {
        wchar_t valueBuffer[64] = {};
        GetWindowTextW(state->editControl, valueBuffer, static_cast<int>(ARRAYSIZE(valueBuffer)));
        wchar_t* parseEnd = nullptr;
        double baseValue = wcstod(valueBuffer, &parseEnd);
        while (parseEnd != nullptr && *parseEnd != L'\0' && iswspace(*parseEnd) != 0) {
          ++parseEnd;
        }
        if (parseEnd == valueBuffer || (parseEnd != nullptr && *parseEnd != L'\0') || !std::isfinite(baseValue)) {
          MessageBoxW(hwnd, L"Invalid value. Enter a numeric interval.", L"Backrest Watcher", MB_ICONWARNING | MB_OK);
          SetFocus(state->editControl);
          return 0;
        }

        const bool useMinutes = (SendMessageW(state->minutesCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
        double intervalSeconds = useMinutes ? (baseValue * 60.0) : baseValue;
        if (!std::isfinite(intervalSeconds) || intervalSeconds < kMinMonitorIntervalSeconds) {
          MessageBoxW(hwnd, L"Value must be at least 0.5 seconds.", L"Backrest Watcher", MB_ICONWARNING | MB_OK);
          SetFocus(state->editControl);
          return 0;
        }
        if (intervalSeconds > kMaxTimerSupportedSeconds) {
          MessageBoxW(hwnd, L"Value exceeds Windows timer technical limits.", L"Backrest Watcher", MB_ICONWARNING | MB_OK);
          SetFocus(state->editControl);
          return 0;
        }

        state->resultValueMs = static_cast<UINT>(std::llround(intervalSeconds * 1000.0));
        state->accepted = true;
        DestroyWindow(hwnd);
        return 0;
      }
      if (LOWORD(wParam) == IDCANCEL) {
        DestroyWindow(hwnd);
        return 0;
      }
      return 0;

    case WM_CLOSE:
      DestroyWindow(hwnd);
      return 0;

    case WM_DESTROY:
      return 0;

    default:
      return DefWindowProcW(hwnd, message, wParam, lParam);
  }
}

bool PromptMonitorIntervalMs(HWND parent, UINT currentValueMs, UINT* outValueMs) {
  const wchar_t kDialogClassName[] = L"BackrestIntervalInputWindowClass";
  static bool registered = false;
  if (!registered) {
    WNDCLASSW cls = {};
    cls.lpfnWndProc = IntervalInputWindowProc;
    cls.hInstance = GetModuleHandleW(nullptr);
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    cls.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    cls.lpszClassName = kDialogClassName;
    RegisterClassW(&cls);
    registered = true;
  }

  IntervalInputDialogState state = {};
  state.initialValueMs = currentValueMs;

  HWND dialog = CreateWindowExW(
      WS_EX_DLGMODALFRAME,
      kDialogClassName,
      L"Check Interval",
      WS_CAPTION | WS_POPUP | WS_SYSMENU,
      CW_USEDEFAULT,
      CW_USEDEFAULT,
      300,
      170,
      parent,
      nullptr,
      GetModuleHandleW(nullptr),
      &state);

  if (!dialog) {
    return false;
  }

  RECT parentRect = {};
  GetWindowRect(parent, &parentRect);
  const int dialogWidth = 300;
  const int dialogHeight = 170;
  const int x = parentRect.left + ((parentRect.right - parentRect.left) - dialogWidth) / 2;
  const int y = parentRect.top + ((parentRect.bottom - parentRect.top) - dialogHeight) / 2;
  SetWindowPos(dialog, nullptr, x, y, dialogWidth, dialogHeight, SWP_NOZORDER | SWP_SHOWWINDOW);

  EnableWindow(parent, FALSE);
  MSG msg = {};
  BOOL getMessageResult = 0;
  while (IsWindow(dialog) && (getMessageResult = GetMessageW(&msg, nullptr, 0, 0)) > 0) {
    if (!IsDialogMessageW(dialog, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }
  const bool receivedQuit = (getMessageResult == 0);
  if (receivedQuit && IsWindow(dialog)) {
    DestroyWindow(dialog);
  }
  EnableWindow(parent, TRUE);
  SetForegroundWindow(parent);
  if (receivedQuit) {
    PostQuitMessage(static_cast<int>(msg.wParam));
    return false;
  }

  if (!state.accepted) {
    return false;
  }

  *outValueMs = state.resultValueMs;
  return true;
}

void PromptAndSetMonitorInterval() {
  UINT intervalMs = g_state.monitorIntervalMs;
  if (!PromptMonitorIntervalMs(g_state.hwnd, g_state.monitorIntervalMs, &intervalMs)) {
    return;
  }
  SetMonitorInterval(intervalMs);
}

void ChooseLogPath() {
  wchar_t filePathBuffer[4096] = {};
  StringCchCopyW(filePathBuffer, ARRAYSIZE(filePathBuffer), g_state.logPath.c_str());

  OPENFILENAMEW ofn = {};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = g_state.hwnd;
  ofn.lpstrFile = filePathBuffer;
  ofn.nMaxFile = static_cast<DWORD>(ARRAYSIZE(filePathBuffer));
  ofn.lpstrFilter = L"Log files (*.log)\0*.log\0All files (*.*)\0*.*\0";
  ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
  ofn.lpstrTitle = L"Select backrest.log path";

  if (!GetOpenFileNameW(&ofn)) {
    return;
  }

  g_state.logPath = filePathBuffer;
  SaveLogPathToConfig(g_state.logPath);
  g_state.acknowledgedOffset = 0;
  SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
  ResetWatcherAndRescan();
}

void ShowTrayContextMenu(HWND hwnd) {
  HMENU menu = CreatePopupMenu();
  if (!menu) {
    return;
  }
  wchar_t intervalMenuText[96] = {};
  StringCchPrintfW(
      intervalMenuText,
      ARRAYSIZE(intervalMenuText),
      L"Set check interval... (current: %.3f s)",
      static_cast<double>(g_state.monitorIntervalMs) / 1000.0);

  AppendMenuW(menu, MF_STRING, kMenuSetLogPath, L"Set log file path...");
  AppendMenuW(menu, MF_STRING, kMenuSetMonitorInterval, intervalMenuText);
  AppendMenuW(menu, MF_STRING, kMenuOpenLogFolder, L"Open log folder");
  AppendMenuW(menu, MF_STRING, kMenuOpenLogFile, L"Open log file");
  AppendMenuW(menu, MF_STRING, kMenuAcknowledgeAlert, L"Acknowledge warning/error");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, kMenuExit, L"Exit");

  POINT cursorPos = {};
  GetCursorPos(&cursorPos);
  SetForegroundWindow(hwnd);
  TrackPopupMenu(menu, TPM_BOTTOMALIGN | TPM_LEFTALIGN | TPM_RIGHTBUTTON, cursorPos.x, cursorPos.y, 0, hwnd, nullptr);
  PostMessageW(hwnd, WM_NULL, 0, 0);
  DestroyMenu(menu);
}

bool InitializeTrayIcon(HWND hwnd) {
  const int smallIconWidth = GetSystemMetrics(SM_CXSMICON);
  const int smallIconHeight = GetSystemMetrics(SM_CYSMICON);
  g_state.normalIcon = static_cast<HICON>(LoadImageW(
      nullptr,
      NormalIconPath().c_str(),
      IMAGE_ICON,
      smallIconWidth,
      smallIconHeight,
      LR_LOADFROMFILE));
  g_state.ownsNormalIcon = (g_state.normalIcon != nullptr);
  if (!g_state.normalIcon) {
    g_state.normalIcon = LoadIconW(nullptr, IDI_INFORMATION);
  }
  g_state.alertIcon = LoadIconW(nullptr, IDI_ERROR);
  if (!g_state.normalIcon || !g_state.alertIcon) {
    return false;
  }

  g_state.trayIcon = {};
  g_state.trayIcon.cbSize = sizeof(g_state.trayIcon);
  g_state.trayIcon.hWnd = hwnd;
  g_state.trayIcon.uID = kTrayIconId;
  g_state.trayIcon.uCallbackMessage = kTrayMessage;
  return AddTrayIcon();
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  if (message == g_state.taskbarCreatedMessage && g_state.taskbarCreatedMessage != 0) {
    AddTrayIcon();
    return 0;
  }

  switch (message) {
    case WM_TIMER:
      if (wParam == kMonitorTimerId) {
        MonitorLogFileOnce();
      } else if (wParam == kBlinkTimerId) {
        if (g_state.hasAlert) {
          g_state.blinkShowAlertIcon = !g_state.blinkShowAlertIcon;
          UpdateTrayIcon();
        } else if (!g_state.blinkShowAlertIcon) {
          g_state.blinkShowAlertIcon = true;
          UpdateTrayIcon();
        }
      }
      return 0;

    case kTrayMessage: {
      // For NOTIFYICON_VERSION_4, the event code is carried in LOWORD(lParam).
      const UINT trayEvent = static_cast<UINT>(LOWORD(lParam));
      if (trayEvent == WM_RBUTTONUP || trayEvent == WM_CONTEXTMENU || trayEvent == NIN_KEYSELECT) {
        ShowTrayContextMenu(hwnd);
      } else if (trayEvent == WM_LBUTTONDBLCLK) {
        AcknowledgeAlert();
      }
      return 0;
    }

    case WM_COMMAND: {
      switch (LOWORD(wParam)) {
        case kMenuSetLogPath:
          ChooseLogPath();
          return 0;
        case kMenuOpenLogFolder:
          OpenLogFolder();
          return 0;
        case kMenuOpenLogFile:
          OpenLogFile();
          return 0;
        case kMenuAcknowledgeAlert:
          AcknowledgeAlert();
          return 0;
        case kMenuSetMonitorInterval:
          PromptAndSetMonitorInterval();
          return 0;
        case kMenuExit:
          DestroyWindow(hwnd);
          return 0;
        default:
          return 0;
      }
    }

    case WM_DESTROY:
      KillTimer(hwnd, kMonitorTimerId);
      KillTimer(hwnd, kBlinkTimerId);
      if (g_state.acknowledgePopupHwnd && IsWindow(g_state.acknowledgePopupHwnd)) {
        DestroyWindow(g_state.acknowledgePopupHwnd);
        g_state.acknowledgePopupHwnd = nullptr;
      }
      if (g_state.trayIconAdded) {
        Shell_NotifyIconW(NIM_DELETE, &g_state.trayIcon);
        g_state.trayIconAdded = false;
      }
      if (g_state.ownsNormalIcon && g_state.normalIcon) {
        DestroyIcon(g_state.normalIcon);
        g_state.normalIcon = nullptr;
        g_state.ownsNormalIcon = false;
      }
      PostQuitMessage(0);
      return 0;

    default:
      return DefWindowProcW(hwnd, message, wParam, lParam);
  }
}

}  // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
  bool alreadyRunning = false;
  if (!AcquireSingleInstanceLock(&alreadyRunning)) {
    MessageBoxW(nullptr, L"Failed to initialize single-instance lock.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    return 1;
  }
  if (alreadyRunning) {
    ReleaseSingleInstanceLock();
    return 0;
  }

  g_state.taskbarCreatedMessage = RegisterWindowMessageW(L"TaskbarCreated");
  g_state.configPath = ConfigFilePath();
  LoadLogPathFromConfig();

  const wchar_t kWindowClassName[] = L"BackrestTrayWatcherWindowClass";
  WNDCLASSEXW windowClass = {};
  windowClass.cbSize = sizeof(windowClass);
  windowClass.lpfnWndProc = WindowProc;
  windowClass.hInstance = instance;
  windowClass.lpszClassName = kWindowClassName;
  windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);

  if (!RegisterClassExW(&windowClass)) {
    MessageBoxW(nullptr, L"Failed to register window class.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    ReleaseSingleInstanceLock();
    return 1;
  }

  HWND hwnd = CreateWindowExW(
      0,
      kWindowClassName,
      L"Backrest Watcher",
      WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT,
      CW_USEDEFAULT,
      CW_USEDEFAULT,
      CW_USEDEFAULT,
      nullptr,
      nullptr,
      instance,
      nullptr);

  if (!hwnd) {
    MessageBoxW(nullptr, L"Failed to create hidden window.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    ReleaseSingleInstanceLock();
    return 1;
  }

  g_state.hwnd = hwnd;

  if (!InitializeTrayIcon(hwnd)) {
    MessageBoxW(hwnd, L"Failed to add tray icon.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    DestroyWindow(hwnd);
    ReleaseSingleInstanceLock();
    return 1;
  }

  if (SetTimer(hwnd, kMonitorTimerId, g_state.monitorIntervalMs, nullptr) == 0) {
    MessageBoxW(hwnd, L"Failed to start log monitoring timer.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    DestroyWindow(hwnd);
    ReleaseSingleInstanceLock();
    return 1;
  }
  if (SetTimer(hwnd, kBlinkTimerId, kBlinkIntervalMs, nullptr) == 0) {
    MessageBoxW(hwnd, L"Failed to start icon blinking timer.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    DestroyWindow(hwnd);
    ReleaseSingleInstanceLock();
    return 1;
  }
  ResetWatcherAndRescan();

  ShowWindow(hwnd, SW_HIDE);
  UpdateWindow(hwnd);

  MSG message = {};
  while (GetMessageW(&message, nullptr, 0, 0) > 0) {
    TranslateMessage(&message);
    DispatchMessageW(&message);
  }

  ReleaseSingleInstanceLock();
  return static_cast<int>(message.wParam);
}
