#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#include <commdlg.h>
#include <strsafe.h>

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <limits>
#include <string>
#include <string_view>
#include <system_error>
#include <utility>
#include <vector>

#pragma comment(lib, "comctl32.lib")

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
constexpr UINT kMenuOpenAlertMessages = 1006;
constexpr UINT kMenuSetMonitorInterval = 1101;
constexpr UINT kMenuSetDoubleClickAction = 1102;
constexpr UINT_PTR kAcknowledgePopupTimerId = 3;
constexpr UINT kDefaultAcknowledgePopupDurationMs = 2500;
constexpr UINT kMinAcknowledgePopupDurationMs = 500;
constexpr UINT kMaxAcknowledgePopupDurationMs = 30000;
constexpr std::string_view kAlertKeyword = "\"logger\":";
constexpr std::string_view kWarnLevelKeyword = "\"level\":\"warn\"";
constexpr std::string_view kErrorLevelKeyword = "\"level\":\"error\"";
constexpr std::string_view kDPanicLevelKeyword = "\"level\":\"dpanic\"";
constexpr std::string_view kPanicLevelKeyword = "\"level\":\"panic\"";
constexpr std::string_view kFatalLevelKeyword = "\"level\":\"fatal\"";
constexpr int kAcknowledgePopupWidth = 420;
constexpr int kAcknowledgePopupHeight = 72;
constexpr int kAcknowledgePopupOffsetPx = 8;
constexpr wchar_t kAcknowledgePopupWindowClassName[] = L"BackrestWatcherAcknowledgePopupWindowClass";
constexpr wchar_t kAlertManagerWindowClassName[] = L"BackrestWatcherAlertManagerWindowClass";
constexpr wchar_t kDoubleClickActionDialogClassName[] = L"BackrestWatcherDoubleClickActionWindowClass";
constexpr wchar_t kIgnoreFileName[] = L"Ignore.txt";
constexpr wchar_t kDebugLogFileName[] = L"debuglog.txt";
constexpr int kAlertManagerWindowWidth = 760;
constexpr int kAlertManagerWindowHeight = 520;
constexpr int kHiddenOwnerWindowWidth = 360;
constexpr int kHiddenOwnerWindowHeight = 220;
constexpr int kAlertManagerPaddingPx = 12;
constexpr int kAlertManagerButtonWidth = 170;
constexpr int kAlertManagerButtonHeight = 28;
constexpr int kAlertManagerButtonSpacingPx = 8;
constexpr int kAlertManagerDetailsHeight = 138;
constexpr std::string_view kIgnoreRuleSeparator = "&&";
constexpr wchar_t kIgnoreUsageHintText[] =
    L"Ignore.txt usage:\r\n"
    L"- Each line is one rule.\r\n"
    L"- Use && to require all parts.\r\n"
    L"- Example rule:\r\n"
    L"\"level\":\"warn\" && \"msg\":\"error processing item\"\r\n"
    L"This rule will ignore any log line that contains both of the specified parts: "
    L"\"level\":\"warn\" and \"msg\":\"error processing item\".\r\n"
    L"- After editing Ignore.txt manually, click Refresh list.";
constexpr UINT kAlertManagerTabMessages = 0;
constexpr UINT kAlertManagerTabIgnored = 1;
constexpr int kControlAlertTab = 2001;
constexpr int kControlActiveAlertsList = 2002;
constexpr int kControlActiveAlertsDetails = 2003;
constexpr int kControlIgnoreSelectedButton = 2004;
constexpr int kControlIgnoredAlertsList = 2005;
constexpr int kControlIgnoredAlertsDetails = 2006;
constexpr int kControlRefreshIgnoredButton = 2007;
constexpr int kControlOpenIgnoreFileButton = 2008;
constexpr int kControlIgnoreUsageHint = 2009;
constexpr int kControlDoubleClickActionCombo = 2101;

enum class AlertSeverity {
  kNone = 0,
  kWarning = 1,
  kError = 2,
};

enum class DoubleClickAction : UINT {
  kOpenAlertMessages = kMenuOpenAlertMessages,
  kSetLogPath = kMenuSetLogPath,
  kSetMonitorInterval = kMenuSetMonitorInterval,
  kOpenLogFolder = kMenuOpenLogFolder,
  kOpenLogFile = kMenuOpenLogFile,
  kAcknowledgeAlert = kMenuAcknowledgeAlert,
};

struct AlertEntry {
  AlertSeverity severity = AlertSeverity::kNone;
  bool isIgnored = false;
  ULONGLONG lineNumber = 0;
  std::string rawLine;
  std::string ignoreRuleText;
  std::string matchedIgnoreRuleText;
  std::wstring listText;
  std::wstring summaryText;
  std::wstring itemText;
  std::wstring errorMessageText;
  std::wstring detailText;
};

struct IgnoreRule {
  std::string text;
  std::vector<std::string> requiredTerms;
};

struct AppState {
  HWND hwnd = nullptr;
  std::wstring configPath;
  std::wstring ignorePath;
  std::wstring debugLogPath;
  std::wstring logPath;
  ULONGLONG acknowledgedOffset = 0;
  ULONGLONG lastOffset = 0;
  ULONGLONG lastLineNumber = 0;
  UINT monitorIntervalMs = kDefaultMonitorIntervalMs;
  bool monitorIntervalUseMinutes = false;
  bool debugMode = false;
  DoubleClickAction doubleClickAction = DoubleClickAction::kOpenLogFile;
  AlertSeverity alertSeverity = AlertSeverity::kNone;
  bool blinkShowAlertIcon = true;
  UINT acknowledgePopupDurationMs = kDefaultAcknowledgePopupDurationMs;
  NOTIFYICONDATAW trayIcon = {};
  HICON normalIcon = nullptr;
  HICON warningIcon = nullptr;
  HICON errorIcon = nullptr;
  bool ownsNormalIcon = false;
  bool trayIconAdded = false;
  HWND acknowledgePopupHwnd = nullptr;
  HWND alertManagerHwnd = nullptr;
  HWND alertManagerTabHwnd = nullptr;
  HWND activeAlertsListHwnd = nullptr;
  HWND activeAlertsDetailsHwnd = nullptr;
  HWND ignoreSelectedButtonHwnd = nullptr;
  HWND ignoredAlertsListHwnd = nullptr;
  HWND ignoredAlertsDetailsHwnd = nullptr;
  HWND refreshIgnoredButtonHwnd = nullptr;
  HWND openIgnoreFileButtonHwnd = nullptr;
  HWND ignoreUsageHintHwnd = nullptr;
  HANDLE singleInstanceMutex = nullptr;
  UINT taskbarCreatedMessage = 0;
  std::vector<IgnoreRule> ignoredRules;
  bool ignoreListStateKnown = false;
  bool ignoreFileExists = false;
  std::filesystem::file_time_type ignoreFileLastWriteTime = {};
  std::vector<AlertEntry> activeAlertEntries;
};

AppState g_state;

struct IntervalInputDialogState {
  UINT initialValueMs = 0;
  bool initialUseMinutes = false;
  UINT resultValueMs = 0;
  bool resultUseMinutes = false;
  bool accepted = false;
  HWND editControl = nullptr;
  HWND secondsCheckbox = nullptr;
  HWND minutesCheckbox = nullptr;
};

struct DoubleClickActionDialogState {
  DoubleClickAction initialAction = DoubleClickAction::kOpenLogFile;
  DoubleClickAction resultAction = DoubleClickAction::kOpenLogFile;
  bool accepted = false;
  HWND comboBox = nullptr;
};

HMENU MenuHandleFromId(UINT id) {
  return reinterpret_cast<HMENU>(static_cast<INT_PTR>(id));
}

void RefreshAlertManagerWindowContent();
void ShowAlertManagerWindow();
void PromptAndSetDoubleClickAction();
bool HandleCommand(UINT commandId);
void ResetWatcherAndRescan();
bool TryBuildAlertEntryFromLine(std::string_view line, AlertEntry* outEntry);
bool ReloadIgnoreListIfChanged(bool forceReload);
void OpenLogFile();
AlertSeverity MaxAlertSeverity(AlertSeverity left, AlertSeverity right);
const IgnoreRule* FindMatchingIgnoreRule(std::string_view rawLine);

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

std::wstring IgnoreFilePath() {
  return ExeDirectory() + L"\\" + kIgnoreFileName;
}

std::wstring DebugLogPath() {
  return ExeDirectory() + L"\\" + kDebugLogFileName;
}

std::string WideToUtf8(std::wstring_view text) {
  if (text.empty()) {
    return {};
  }

  const int requiredChars = WideCharToMultiByte(
      CP_UTF8,
      0,
      text.data(),
      static_cast<int>(text.size()),
      nullptr,
      0,
      nullptr,
      nullptr);
  if (requiredChars <= 0) {
    std::string fallback;
    fallback.reserve(text.size());
    for (const wchar_t ch : text) {
      fallback.push_back(static_cast<char>(ch & 0xFF));
    }
    return fallback;
  }

  std::string utf8(requiredChars, '\0');
  WideCharToMultiByte(
      CP_UTF8,
      0,
      text.data(),
      static_cast<int>(text.size()),
      utf8.data(),
      requiredChars,
      nullptr,
      nullptr);
  return utf8;
}

bool EqualsTextInsensitive(std::wstring_view left, std::wstring_view right) {
  return CompareStringOrdinal(
             left.data(),
             static_cast<int>(left.size()),
             right.data(),
             static_cast<int>(right.size()),
             TRUE) == CSTR_EQUAL;
}

bool IsDebugArgument(std::wstring_view argument) {
  return EqualsTextInsensitive(argument, L"debug") ||
         EqualsTextInsensitive(argument, L"--debug") ||
         EqualsTextInsensitive(argument, L"/debug");
}

std::wstring ReadEnvironmentVariableValue(const wchar_t* variableName) {
  if (!variableName || !*variableName) {
    return {};
  }

  const DWORD requiredChars = GetEnvironmentVariableW(variableName, nullptr, 0);
  if (requiredChars == 0) {
    return {};
  }

  std::wstring value(requiredChars, L'\0');
  const DWORD writtenChars =
      GetEnvironmentVariableW(variableName, value.data(), requiredChars);
  if (writtenChars == 0 || writtenChars >= requiredChars) {
    return {};
  }
  value.resize(writtenChars);
  return value;
}

std::wstring WindowMessageName(UINT message) {
  switch (message) {
    case WM_CREATE:
      return L"WM_CREATE";
    case WM_DESTROY:
      return L"WM_DESTROY";
    case WM_CLOSE:
      return L"WM_CLOSE";
    case WM_COMMAND:
      return L"WM_COMMAND";
    case WM_TIMER:
      return L"WM_TIMER";
    case WM_NOTIFY:
      return L"WM_NOTIFY";
    case WM_SIZE:
      return L"WM_SIZE";
    case WM_PAINT:
      return L"WM_PAINT";
    case WM_SETFOCUS:
      return L"WM_SETFOCUS";
    case WM_LBUTTONDOWN:
      return L"WM_LBUTTONDOWN";
    case WM_LBUTTONUP:
      return L"WM_LBUTTONUP";
    case WM_LBUTTONDBLCLK:
      return L"WM_LBUTTONDBLCLK";
    case WM_RBUTTONDOWN:
      return L"WM_RBUTTONDOWN";
    case WM_RBUTTONUP:
      return L"WM_RBUTTONUP";
    case WM_CONTEXTMENU:
      return L"WM_CONTEXTMENU";
    case WM_NULL:
      return L"WM_NULL";
    case kTrayMessage:
      return L"kTrayMessage";
    default:
      return L"MSG_" + std::to_wstring(message);
  }
}

void DebugLog(std::wstring_view message) {
  if (!g_state.debugMode || g_state.debugLogPath.empty()) {
    return;
  }

  SYSTEMTIME localTime = {};
  GetLocalTime(&localTime);

  wchar_t prefix[64] = {};
  StringCchPrintfW(
      prefix,
      ARRAYSIZE(prefix),
      L"[%04u-%02u-%02u %02u:%02u:%02u.%03u] ",
      localTime.wYear,
      localTime.wMonth,
      localTime.wDay,
      localTime.wHour,
      localTime.wMinute,
      localTime.wSecond,
      localTime.wMilliseconds);

  std::wstring line(prefix);
  line.append(message);
  line.append(L"\r\n");

  std::ofstream outputFile(std::filesystem::path(g_state.debugLogPath), std::ios::binary | std::ios::app);
  if (!outputFile.is_open()) {
    return;
  }

  const std::string utf8 = WideToUtf8(line);
  outputFile.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
}

void DebugLogWindowMessage(std::wstring_view windowName, UINT message, WPARAM wParam, LPARAM lParam) {
  DebugLog(
      std::wstring(windowName) +
      L" received " +
      WindowMessageName(message) +
      L" wParam=" + std::to_wstring(static_cast<unsigned long long>(wParam)) +
      L" lParam=" + std::to_wstring(static_cast<long long>(lParam)));
}

void InitializeDebugModeFromCommandLine() {
  g_state.debugLogPath = DebugLogPath();
  g_state.debugMode = false;

  int argc = 0;
  LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
  if (argv) {
    for (int index = 1; index < argc; ++index) {
      if (IsDebugArgument(argv[index])) {
        g_state.debugMode = true;
        break;
      }
    }
    LocalFree(argv);
  }

  if (g_state.debugMode) {
    DebugLog(L"===== Debug session started =====");
    DebugLog(std::wstring(L"Command line: ") + GetCommandLineW());
  }
}

RECT CursorMonitorWorkArea() {
  POINT cursorPos = {};
  if (!GetCursorPos(&cursorPos)) {
    cursorPos.x = GetSystemMetrics(SM_CXSCREEN) / 2;
    cursorPos.y = GetSystemMetrics(SM_CYSCREEN) / 2;
  }

  RECT workArea = {0, 0, GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)};
  MONITORINFO monitorInfo = {};
  monitorInfo.cbSize = sizeof(monitorInfo);
  HMONITOR monitor = MonitorFromPoint(cursorPos, MONITOR_DEFAULTTONEAREST);
  if (monitor && GetMonitorInfoW(monitor, &monitorInfo)) {
    workArea = monitorInfo.rcWork;
  }
  return workArea;
}

POINT CenteredWindowPosition(int width, int height) {
  const RECT workArea = CursorMonitorWorkArea();
  const int workWidth = static_cast<int>(workArea.right - workArea.left);
  const int workHeight = static_cast<int>(workArea.bottom - workArea.top);

  POINT position = {};
  position.x = workArea.left;
  position.y = workArea.top;

  if (width < workWidth) {
    position.x = workArea.left + ((workWidth - width) / 2);
  }
  if (height < workHeight) {
    position.y = workArea.top + ((workHeight - height) / 2);
  }

  return position;
}

void CenterWindowOnCurrentScreen(HWND hwnd, int width, int height, UINT flags) {
  const POINT position = CenteredWindowPosition(width, height);
  SetWindowPos(hwnd, nullptr, position.x, position.y, width, height, flags);
}

const wchar_t* AlertSeverityLabel(AlertSeverity severity) {
  switch (severity) {
    case AlertSeverity::kWarning:
      return L"WARN";
    case AlertSeverity::kError:
      return L"ERROR";
    case AlertSeverity::kNone:
    default:
      return L"OK";
  }
}

const wchar_t* DoubleClickActionLabel(DoubleClickAction action) {
  switch (action) {
    case DoubleClickAction::kOpenAlertMessages:
      return L"Open alert messages";
    case DoubleClickAction::kSetLogPath:
      return L"Set log file path";
    case DoubleClickAction::kSetMonitorInterval:
      return L"Set check interval";
    case DoubleClickAction::kOpenLogFolder:
      return L"Open log folder";
    case DoubleClickAction::kOpenLogFile:
      return L"Open log file";
    case DoubleClickAction::kAcknowledgeAlert:
      return L"Acknowledge warning/error";
    default:
      return L"Open log file";
  }
}

const wchar_t* DoubleClickActionConfigValue(DoubleClickAction action) {
  switch (action) {
    case DoubleClickAction::kOpenAlertMessages:
      return L"open_alert_messages";
    case DoubleClickAction::kSetLogPath:
      return L"set_log_path";
    case DoubleClickAction::kSetMonitorInterval:
      return L"set_monitor_interval";
    case DoubleClickAction::kOpenLogFolder:
      return L"open_log_folder";
    case DoubleClickAction::kOpenLogFile:
      return L"open_log_file";
    case DoubleClickAction::kAcknowledgeAlert:
      return L"acknowledge_alert";
    default:
      return L"open_log_file";
  }
}

bool TryParseDoubleClickAction(const wchar_t* value, DoubleClickAction* outAction) {
  if (!value || !outAction) {
    return false;
  }
  if (lstrcmpiW(value, L"open_alert_messages") == 0) {
    *outAction = DoubleClickAction::kOpenAlertMessages;
    return true;
  }
  if (lstrcmpiW(value, L"set_log_path") == 0) {
    *outAction = DoubleClickAction::kSetLogPath;
    return true;
  }
  if (lstrcmpiW(value, L"set_monitor_interval") == 0) {
    *outAction = DoubleClickAction::kSetMonitorInterval;
    return true;
  }
  if (lstrcmpiW(value, L"open_log_folder") == 0) {
    *outAction = DoubleClickAction::kOpenLogFolder;
    return true;
  }
  if (lstrcmpiW(value, L"open_log_file") == 0) {
    *outAction = DoubleClickAction::kOpenLogFile;
    return true;
  }
  if (lstrcmpiW(value, L"acknowledge_alert") == 0) {
    *outAction = DoubleClickAction::kAcknowledgeAlert;
    return true;
  }
  return false;
}

constexpr DoubleClickAction kDoubleClickActionOptions[] = {
    DoubleClickAction::kOpenAlertMessages,
    DoubleClickAction::kOpenLogFile,
    DoubleClickAction::kOpenLogFolder,
    DoubleClickAction::kAcknowledgeAlert,
    DoubleClickAction::kSetLogPath,
    DoubleClickAction::kSetMonitorInterval,
};

std::wstring Utf8ToWide(std::string_view text) {
  if (text.empty()) {
    return {};
  }

  const int requiredChars = MultiByteToWideChar(
      CP_UTF8,
      0,
      text.data(),
      static_cast<int>(text.size()),
      nullptr,
      0);
  if (requiredChars <= 0) {
    std::wstring fallback;
    fallback.reserve(text.size());
    for (const unsigned char ch : text) {
      fallback.push_back(static_cast<wchar_t>(ch));
    }
    return fallback;
  }

  std::wstring wide(requiredChars, L'\0');
  MultiByteToWideChar(
      CP_UTF8,
      0,
      text.data(),
      static_cast<int>(text.size()),
      wide.data(),
      requiredChars);
  return wide;
}

std::string_view TrimAsciiWhitespace(std::string_view text) {
  size_t begin = 0;
  while (begin < text.size() && std::isspace(static_cast<unsigned char>(text[begin])) != 0) {
    ++begin;
  }

  size_t end = text.size();
  while (end > begin && std::isspace(static_cast<unsigned char>(text[end - 1])) != 0) {
    --end;
  }

  return text.substr(begin, end - begin);
}

void AppendUtf8CodePoint(unsigned int codePoint, std::string* output) {
  if (!output) {
    return;
  }
  if (codePoint <= 0x7F) {
    output->push_back(static_cast<char>(codePoint));
  } else if (codePoint <= 0x7FF) {
    output->push_back(static_cast<char>(0xC0 | ((codePoint >> 6) & 0x1F)));
    output->push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
  } else if (codePoint <= 0xFFFF) {
    output->push_back(static_cast<char>(0xE0 | ((codePoint >> 12) & 0x0F)));
    output->push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
    output->push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
  } else {
    output->push_back(static_cast<char>(0xF0 | ((codePoint >> 18) & 0x07)));
    output->push_back(static_cast<char>(0x80 | ((codePoint >> 12) & 0x3F)));
    output->push_back(static_cast<char>(0x80 | ((codePoint >> 6) & 0x3F)));
    output->push_back(static_cast<char>(0x80 | (codePoint & 0x3F)));
  }
}

bool ParseHexDigit(char ch, unsigned int* outValue) {
  if (!outValue) {
    return false;
  }
  if (ch >= '0' && ch <= '9') {
    *outValue = static_cast<unsigned int>(ch - '0');
    return true;
  }
  if (ch >= 'a' && ch <= 'f') {
    *outValue = static_cast<unsigned int>(10 + (ch - 'a'));
    return true;
  }
  if (ch >= 'A' && ch <= 'F') {
    *outValue = static_cast<unsigned int>(10 + (ch - 'A'));
    return true;
  }
  return false;
}

bool ParseJsonStringLiteral(
    std::string_view text,
    size_t openingQuotePos,
    size_t* outNextPos,
    std::string* outValue) {
  if (!outValue || openingQuotePos >= text.size() || text[openingQuotePos] != '"') {
    return false;
  }

  outValue->clear();
  for (size_t pos = openingQuotePos + 1; pos < text.size(); ++pos) {
    const char ch = text[pos];
    if (ch == '"') {
      if (outNextPos) {
        *outNextPos = pos + 1;
      }
      return true;
    }
    if (ch != '\\') {
      outValue->push_back(ch);
      continue;
    }

    ++pos;
    if (pos >= text.size()) {
      return false;
    }

    const char escaped = text[pos];
    switch (escaped) {
      case '"':
      case '\\':
      case '/':
        outValue->push_back(escaped);
        break;
      case 'b':
        outValue->push_back('\b');
        break;
      case 'f':
        outValue->push_back('\f');
        break;
      case 'n':
        outValue->push_back('\n');
        break;
      case 'r':
        outValue->push_back('\r');
        break;
      case 't':
        outValue->push_back('\t');
        break;
      case 'u': {
        if (pos + 4 >= text.size()) {
          return false;
        }
        unsigned int codePoint = 0;
        for (size_t i = 0; i < 4; ++i) {
          unsigned int nibble = 0;
          if (!ParseHexDigit(text[pos + 1 + i], &nibble)) {
            return false;
          }
          codePoint = (codePoint << 4) | nibble;
        }
        AppendUtf8CodePoint(codePoint, outValue);
        pos += 4;
        break;
      }
      default:
        outValue->push_back(escaped);
        break;
    }
  }

  return false;
}

bool ExtractJsonStringField(std::string_view json, std::string_view fieldName, std::string* outValue) {
  if (!outValue || fieldName.empty()) {
    return false;
  }

  const std::string pattern = "\"" + std::string(fieldName) + "\"";
  size_t searchPos = 0;
  while (searchPos < json.size()) {
    const size_t fieldPos = json.find(pattern, searchPos);
    if (fieldPos == std::string_view::npos) {
      return false;
    }

    size_t valuePos = fieldPos + pattern.size();
    while (valuePos < json.size() && std::isspace(static_cast<unsigned char>(json[valuePos])) != 0) {
      ++valuePos;
    }
    if (valuePos >= json.size() || json[valuePos] != ':') {
      searchPos = fieldPos + pattern.size();
      continue;
    }

    ++valuePos;
    while (valuePos < json.size() && std::isspace(static_cast<unsigned char>(json[valuePos])) != 0) {
      ++valuePos;
    }
    if (valuePos >= json.size() || json[valuePos] != '"') {
      searchPos = fieldPos + pattern.size();
      continue;
    }

    return ParseJsonStringLiteral(json, valuePos, nullptr, outValue);
  }

  return false;
}

bool ExtractErrorMessageField(std::string_view line, std::string* outValue) {
  if (!outValue) {
    return false;
  }

  const size_t errorFieldPos = line.find("\"error\"");
  if (errorFieldPos == std::string_view::npos) {
    return false;
  }
  return ExtractJsonStringField(line.substr(errorFieldPos), "message", outValue);
}

bool TryFindJsonFieldRange(
    std::string_view json,
    std::string_view fieldName,
    size_t* outFieldPos,
    size_t* outFieldEnd) {
  if (!outFieldPos || !outFieldEnd || fieldName.empty()) {
    return false;
  }

  const std::string pattern = "\"" + std::string(fieldName) + "\"";
  size_t searchPos = 0;
  while (searchPos < json.size()) {
    const size_t fieldPos = json.find(pattern, searchPos);
    if (fieldPos == std::string_view::npos) {
      return false;
    }

    size_t valuePos = fieldPos + pattern.size();
    while (valuePos < json.size() && std::isspace(static_cast<unsigned char>(json[valuePos])) != 0) {
      ++valuePos;
    }
    if (valuePos >= json.size() || json[valuePos] != ':') {
      searchPos = fieldPos + pattern.size();
      continue;
    }

    ++valuePos;
    while (valuePos < json.size() && std::isspace(static_cast<unsigned char>(json[valuePos])) != 0) {
      ++valuePos;
    }
    if (valuePos >= json.size()) {
      return false;
    }

    size_t fieldEnd = valuePos;
    if (json[valuePos] == '"') {
      std::string ignoredValue;
      if (!ParseJsonStringLiteral(json, valuePos, &fieldEnd, &ignoredValue)) {
        return false;
      }
    } else if (json[valuePos] == '{' || json[valuePos] == '[') {
      const char opening = json[valuePos];
      const char closing = (opening == '{') ? '}' : ']';
      int depth = 0;
      bool inString = false;
      bool escaping = false;
      for (fieldEnd = valuePos; fieldEnd < json.size(); ++fieldEnd) {
        const char ch = json[fieldEnd];
        if (inString) {
          if (escaping) {
            escaping = false;
          } else if (ch == '\\') {
            escaping = true;
          } else if (ch == '"') {
            inString = false;
          }
          continue;
        }

        if (ch == '"') {
          inString = true;
          continue;
        }
        if (ch == opening) {
          ++depth;
        } else if (ch == closing) {
          --depth;
          if (depth == 0) {
            ++fieldEnd;
            break;
          }
        }
      }
      if (fieldEnd > json.size()) {
        fieldEnd = json.size();
      }
    } else {
      while (fieldEnd < json.size() && json[fieldEnd] != ',' && json[fieldEnd] != '}' && json[fieldEnd] != ']') {
        ++fieldEnd;
      }
      while (fieldEnd > valuePos &&
             std::isspace(static_cast<unsigned char>(json[fieldEnd - 1])) != 0) {
        --fieldEnd;
      }
    }

    *outFieldPos = fieldPos;
    *outFieldEnd = fieldEnd;
    return true;
  }

  return false;
}

std::wstring FormatAlertSummaryText(
    AlertSeverity severity,
    std::wstring_view loggerText,
    std::wstring_view messageText,
    std::wstring_view fallbackRawLine) {
  std::wstring summary = L"[";
  summary += AlertSeverityLabel(severity);
  summary += L"] ";

  if (!loggerText.empty()) {
    summary += loggerText;
  }
  if (!loggerText.empty() && !messageText.empty()) {
    summary += L" | ";
  }
  if (!messageText.empty()) {
    summary += messageText;
  }
  if (loggerText.empty() && messageText.empty()) {
    summary += fallbackRawLine;
  }
  return summary;
}

std::wstring FormatAlertListText(bool isIgnored, std::wstring_view rawLineText) {
  if (!isIgnored) {
    return std::wstring(rawLineText);
  }

  std::wstring listText = L"(Ignored) ";
  listText += rawLineText;
  return listText;
}

std::wstring FormatAlertDetailText(const AlertEntry& entry) {
  std::wstring details = L"Line: ";
  details += std::to_wstring(entry.lineNumber);
  details += L"\r\nStatus: ";
  details += entry.isIgnored ? L"Ignored" : L"Active";
  details += L"\r\nSummary: ";
  details += entry.summaryText;
  if (!entry.matchedIgnoreRuleText.empty()) {
    details += L"\r\nMatched ignore rule: ";
    details += Utf8ToWide(entry.matchedIgnoreRuleText);
  }
  if (!entry.ignoreRuleText.empty()) {
    details += L"\r\nSuggested ignore rule: ";
    details += Utf8ToWide(entry.ignoreRuleText);
  }
  if (!entry.itemText.empty()) {
    details += L"\r\nItem: ";
    details += entry.itemText;
  }
  if (!entry.errorMessageText.empty()) {
    details += L"\r\nError: ";
    details += entry.errorMessageText;
  }
  details += L"\r\nRaw:\r\n";
  details += Utf8ToWide(entry.rawLine);
  if (entry.isIgnored) {
    details += L"\r\n\r\nIgnored messages still appear here, but they do not affect the tray icon severity.";
  }
  return details;
}

void UpdateAlertEntryPresentation(AlertEntry* entry) {
  if (!entry) {
    return;
  }

  const std::wstring rawLineText = Utf8ToWide(entry->rawLine);
  entry->listText = FormatAlertListText(entry->isIgnored, rawLineText);
  entry->detailText = FormatAlertDetailText(*entry);
}

ULONGLONG CountLogicalLinesUpToOffset(HANDLE file, ULONGLONG endOffset) {
  if (file == INVALID_HANDLE_VALUE || endOffset == 0) {
    return 0;
  }

  LARGE_INTEGER filePointer = {};
  filePointer.QuadPart = 0;
  if (!SetFilePointerEx(file, filePointer, nullptr, FILE_BEGIN)) {
    return 0;
  }

  constexpr DWORD kBufferSize = 64 * 1024;
  char buffer[kBufferSize];
  ULONGLONG remaining = endOffset;
  ULONGLONG lineCount = 0;
  bool sawAnyByte = false;
  char lastByte = '\n';
  while (remaining > 0) {
    const DWORD toRead = static_cast<DWORD>(
        std::min<ULONGLONG>(remaining, static_cast<ULONGLONG>(kBufferSize)));
    DWORD bytesRead = 0;
    if (!ReadFile(file, buffer, toRead, &bytesRead, nullptr) || bytesRead == 0) {
      break;
    }

    sawAnyByte = true;
    lineCount += static_cast<ULONGLONG>(std::count(buffer, buffer + bytesRead, '\n'));
    lastByte = buffer[bytesRead - 1];
    remaining -= bytesRead;
  }

  if (sawAnyByte && lastByte != '\n') {
    ++lineCount;
  }
  return lineCount;
}

bool DoesOffsetStartOnNewLine(HANDLE file, ULONGLONG offset) {
  if (file == INVALID_HANDLE_VALUE || offset == 0) {
    return true;
  }

  LARGE_INTEGER filePointer = {};
  filePointer.QuadPart = static_cast<LONGLONG>(offset - 1);
  if (!SetFilePointerEx(file, filePointer, nullptr, FILE_BEGIN)) {
    return true;
  }

  char previousByte = '\n';
  DWORD bytesRead = 0;
  if (!ReadFile(file, &previousByte, 1, &bytesRead, nullptr) || bytesRead != 1) {
    return true;
  }
  return previousByte == '\n';
}

ULONGLONG StartingLineNumberForOffset(HANDLE file, ULONGLONG offset) {
  ULONGLONG lineNumber = CountLogicalLinesUpToOffset(file, offset);
  if (offset > 0 && !DoesOffsetStartOnNewLine(file, offset) && lineNumber > 0) {
    --lineNumber;
  }
  return lineNumber;
}

void AppendAlertEntryIfNeeded(
    std::string_view line,
    ULONGLONG lineNumber,
    AlertSeverity* inOutHighestSeverity,
    std::vector<AlertEntry>* outEntries) {
  AlertEntry entry = {};
  if (!TryBuildAlertEntryFromLine(line, &entry)) {
    return;
  }

  entry.lineNumber = lineNumber;
  const IgnoreRule* matchedRule = FindMatchingIgnoreRule(entry.rawLine);
  if (matchedRule) {
    entry.isIgnored = true;
    entry.matchedIgnoreRuleText = matchedRule->text;
  }
  UpdateAlertEntryPresentation(&entry);

  if (entry.isIgnored) {
    DebugLog(L"Ignored alert while scanning: " + entry.summaryText);
  } else {
    if (inOutHighestSeverity) {
      *inOutHighestSeverity = MaxAlertSeverity(*inOutHighestSeverity, entry.severity);
    }
    DebugLog(L"Detected alert while scanning: " + entry.summaryText);
  }

  if (outEntries) {
    outEntries->push_back(std::move(entry));
  }
}

void UpdateListBoxHorizontalExtent(HWND listBoxHwnd) {
  if (!listBoxHwnd || !IsWindow(listBoxHwnd)) {
    return;
  }

  const int itemCount = static_cast<int>(SendMessageW(listBoxHwnd, LB_GETCOUNT, 0, 0));
  if (itemCount == LB_ERR || itemCount <= 0) {
    SendMessageW(listBoxHwnd, LB_SETHORIZONTALEXTENT, 0, 0);
    return;
  }

  HDC dc = GetDC(listBoxHwnd);
  if (!dc) {
    return;
  }

  const HFONT font = reinterpret_cast<HFONT>(SendMessageW(listBoxHwnd, WM_GETFONT, 0, 0));
  const HGDIOBJ oldFont = font ? SelectObject(dc, font) : nullptr;
  int maxWidth = 0;
  for (int i = 0; i < itemCount; ++i) {
    const LRESULT textLength = SendMessageW(listBoxHwnd, LB_GETTEXTLEN, static_cast<WPARAM>(i), 0);
    if (textLength == LB_ERR) {
      continue;
    }

    std::wstring text(static_cast<size_t>(textLength), L'\0');
    SendMessageW(listBoxHwnd, LB_GETTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(text.data()));
    SIZE textSize = {};
    if (GetTextExtentPoint32W(dc, text.c_str(), static_cast<int>(text.size()), &textSize)) {
      maxWidth = (std::max)(maxWidth, static_cast<int>(textSize.cx));
    }
  }

  if (oldFont) {
    SelectObject(dc, oldFont);
  }
  ReleaseDC(listBoxHwnd, dc);
  SendMessageW(listBoxHwnd, LB_SETHORIZONTALEXTENT, static_cast<WPARAM>(maxWidth + 24), 0);
}

std::wstring FindNotepadPlusPlusPath() {
  const DWORD bufferSize = SearchPathW(nullptr, L"notepad++.exe", nullptr, 0, nullptr, nullptr);
  if (bufferSize > 0) {
    std::vector<wchar_t> buffer(bufferSize);
    if (SearchPathW(nullptr, L"notepad++.exe", nullptr, bufferSize, buffer.data(), nullptr) > 0) {
      return std::wstring(buffer.data());
    }
  }

  const wchar_t* knownPaths[] = {
      L"D:\\PROGRAMS_D\\Notepad++\\notepad++.exe",
      L"C:\\Program Files\\Notepad++\\notepad++.exe",
      L"C:\\Program Files (x86)\\Notepad++\\notepad++.exe",
  };
  for (const wchar_t* path : knownPaths) {
    const DWORD attributes = GetFileAttributesW(path);
    if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
      return path;
    }
  }

  return L"";
}

void OpenLogFileAtLine(ULONGLONG lineNumber) {
  const std::wstring notepadPlusPlusPath = FindNotepadPlusPlusPath();
  if (!notepadPlusPlusPath.empty()) {
    std::wstring args = L"-nosession -n";
    args += std::to_wstring((std::max)(lineNumber, static_cast<ULONGLONG>(1)));
    args += L" \"";
    args += g_state.logPath;
    args += L"\"";
    const HINSTANCE result = ShellExecuteW(
        nullptr,
        L"open",
        notepadPlusPlusPath.c_str(),
        args.c_str(),
        nullptr,
        SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(result) > 32) {
      return;
    }
  }

  OpenLogFile();
}

bool TryBuildIgnoreRule(std::string_view line, IgnoreRule* outRule) {
  if (!outRule) {
    return false;
  }

  const std::string_view trimmedLine = TrimAsciiWhitespace(line);
  if (trimmedLine.empty()) {
    return false;
  }

  IgnoreRule rule = {};
  rule.text = std::string(trimmedLine);

  std::string_view parseText = trimmedLine;
  size_t termStart = 0;
  while (termStart <= parseText.size()) {
    const size_t separatorPos = parseText.find(kIgnoreRuleSeparator, termStart);
    const size_t termLength =
        (separatorPos == std::string_view::npos) ? (parseText.size() - termStart) : (separatorPos - termStart);
    const std::string_view term = TrimAsciiWhitespace(parseText.substr(termStart, termLength));
    if (!term.empty()) {
      rule.requiredTerms.emplace_back(term);
    }
    if (separatorPos == std::string_view::npos) {
      break;
    }
    termStart = separatorPos + kIgnoreRuleSeparator.size();
  }

  if (rule.requiredTerms.empty()) {
    return false;
  }

  *outRule = std::move(rule);
  return true;
}

std::string SerializeIgnoreRule(const IgnoreRule& rule) {
  return rule.text;
}

std::wstring IgnoreRuleDisplayText(const IgnoreRule& rule) {
  return Utf8ToWide(rule.text);
}

std::wstring IgnoreRuleDetailsText(const IgnoreRule& rule) {
  std::wstring details = (rule.requiredTerms.size() > 1)
      ? L"Type: Contains all parts joined by &&"
      : L"Type: Contains this part";
  details += L"\r\nRule: ";
  details += Utf8ToWide(rule.text);
  details += L"\r\n\r\nRequired parts:";
  for (const std::string& term : rule.requiredTerms) {
    details += L"\r\n- ";
    details += Utf8ToWide(term);
  }
  details += L"\r\n\r\nA log line is ignored only when it contains every part above.";
  return details;
}

bool IsSameIgnoreRule(const IgnoreRule& left, const IgnoreRule& right) {
  return left.text == right.text;
}

bool DoesIgnoreRuleMatchLine(const IgnoreRule& rule, std::string_view rawLine) {
  if (rule.requiredTerms.empty()) {
    return false;
  }

  return std::all_of(
      rule.requiredTerms.begin(),
      rule.requiredTerms.end(),
      [rawLine](const std::string& term) {
        return rawLine.find(term) != std::string_view::npos;
      });
}

const IgnoreRule* FindMatchingIgnoreRule(std::string_view rawLine) {
  const auto it = std::find_if(
      g_state.ignoredRules.begin(),
      g_state.ignoredRules.end(),
      [rawLine](const IgnoreRule& rule) {
        return DoesIgnoreRuleMatchLine(rule, rawLine);
      });
  if (it == g_state.ignoredRules.end()) {
    return nullptr;
  }
  return &(*it);
}

std::string BuildSuggestedIgnoreRuleText(std::string_view line) {
  const std::string_view trimmedLine = TrimAsciiWhitespace(line);
  size_t tsFieldPos = 0;
  size_t tsFieldEnd = 0;
  if (!TryFindJsonFieldRange(trimmedLine, "ts", &tsFieldPos, &tsFieldEnd)) {
    return std::string(trimmedLine);
  }

  const std::string_view left = TrimAsciiWhitespace(trimmedLine.substr(0, tsFieldPos));
  const std::string_view right = TrimAsciiWhitespace(trimmedLine.substr(tsFieldEnd));
  if (left.empty()) {
    return std::string(right);
  }
  if (right.empty()) {
    return std::string(left);
  }

  return std::string(left) + " && " + std::string(right);
}

AlertSeverity MaxAlertSeverity(AlertSeverity left, AlertSeverity right) {
  return (static_cast<int>(left) >= static_cast<int>(right)) ? left : right;
}

bool HasAlert(AlertSeverity severity) {
  return severity != AlertSeverity::kNone;
}

bool ShouldBlinkForSeverity(AlertSeverity severity) {
  return severity == AlertSeverity::kError;
}

AlertSeverity AlertSeverityFromLine(std::string_view line) {
  if (line.find(kAlertKeyword) == std::string_view::npos) {
    return AlertSeverity::kNone;
  }
  if (line.find(kErrorLevelKeyword) != std::string_view::npos ||
      line.find(kDPanicLevelKeyword) != std::string_view::npos ||
      line.find(kPanicLevelKeyword) != std::string_view::npos ||
      line.find(kFatalLevelKeyword) != std::string_view::npos) {
    return AlertSeverity::kError;
  }
  if (line.find(kWarnLevelKeyword) != std::string_view::npos) {
    return AlertSeverity::kWarning;
  }
  return AlertSeverity::kWarning;
}

bool TryBuildAlertEntryFromLine(std::string_view line, AlertEntry* outEntry) {
  if (!outEntry) {
    return false;
  }

  const AlertSeverity severity = AlertSeverityFromLine(line);
  if (!HasAlert(severity)) {
    return false;
  }

  std::string loggerText;
  std::string messageText;
  std::string itemText;
  std::string errorMessageText;
  ExtractJsonStringField(line, "logger", &loggerText);
  ExtractJsonStringField(line, "msg", &messageText);
  ExtractJsonStringField(line, "item", &itemText);
  ExtractErrorMessageField(line, &errorMessageText);

  const std::wstring rawLineText = Utf8ToWide(line);
  const std::wstring loggerWide = Utf8ToWide(loggerText);
  const std::wstring messageWide = Utf8ToWide(messageText);
  const std::wstring itemWide = Utf8ToWide(itemText);
  const std::wstring errorMessageWide = Utf8ToWide(errorMessageText);

  AlertEntry entry = {};
  entry.severity = severity;
  entry.rawLine = std::string(line);
  entry.ignoreRuleText = BuildSuggestedIgnoreRuleText(line);
  entry.summaryText = FormatAlertSummaryText(severity, loggerWide, messageWide, rawLineText);
  entry.itemText = itemWide;
  entry.errorMessageText = errorMessageWide;
  *outEntry = std::move(entry);
  return true;
}

bool IsAlertIgnored(std::string_view rawLine) {
  return FindMatchingIgnoreRule(rawLine) != nullptr;
}

AlertSeverity ScanFileRangeForAlertEntries(
    HANDLE file,
    ULONGLONG beginOffset,
    ULONGLONG endOffset,
    ULONGLONG startingLineNumber,
    ULONGLONG* outEndingLineNumber,
    std::vector<AlertEntry>* outEntries) {
  if (endOffset <= beginOffset) {
    if (outEndingLineNumber) {
      *outEndingLineNumber = startingLineNumber;
    }
    return AlertSeverity::kNone;
  }

  LARGE_INTEGER filePointer = {};
  filePointer.QuadPart = static_cast<LONGLONG>(beginOffset);
  if (!SetFilePointerEx(file, filePointer, nullptr, FILE_BEGIN)) {
    return AlertSeverity::kNone;
  }

  constexpr DWORD kBufferSize = 64 * 1024;
  char buffer[kBufferSize];
  std::string pendingLine;
  AlertSeverity highestSeverity = AlertSeverity::kNone;
  ULONGLONG currentLineNumber = startingLineNumber;

  ULONGLONG remaining = endOffset - beginOffset;
  while (remaining > 0) {
    const DWORD toRead = static_cast<DWORD>(
        std::min<ULONGLONG>(remaining, static_cast<ULONGLONG>(kBufferSize)));
    DWORD bytesRead = 0;
    if (!ReadFile(file, buffer, toRead, &bytesRead, nullptr)) {
      return highestSeverity;
    }
    if (bytesRead == 0) {
      break;
    }

    pendingLine.append(buffer, bytesRead);

    size_t lineStart = 0;
    while (lineStart < pendingLine.size()) {
      const size_t lineEnd = pendingLine.find('\n', lineStart);
      if (lineEnd == std::string::npos) {
        break;
      }

      std::string_view line(pendingLine.data() + lineStart, lineEnd - lineStart);
      if (!line.empty() && line.back() == '\r') {
        line.remove_suffix(1);
      }

      ++currentLineNumber;
      AppendAlertEntryIfNeeded(line, currentLineNumber, &highestSeverity, outEntries);

      lineStart = lineEnd + 1;
    }

    if (lineStart > 0) {
      pendingLine.erase(0, lineStart);
    }

    remaining -= bytesRead;
  }

  if (!pendingLine.empty()) {
    std::string_view line(pendingLine);
    if (!line.empty() && line.back() == '\r') {
      line.remove_suffix(1);
    }
    ++currentLineNumber;
    AppendAlertEntryIfNeeded(line, currentLineNumber, &highestSeverity, outEntries);
  }

  if (outEndingLineNumber) {
    *outEndingLineNumber = currentLineNumber;
  }

  return highestSeverity;
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

bool IsWholeMinutesInterval(UINT intervalMs) {
  return intervalMs >= 60000 && (intervalMs % 60000) == 0;
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

void SaveMonitorIntervalUnitToConfig(bool useMinutes) {
  WritePrivateProfileStringW(
      L"watcher",
      L"monitor_interval_unit",
      useMinutes ? L"minutes" : L"seconds",
      g_state.configPath.c_str());
}

void LoadMonitorIntervalFromConfig() {
  const UINT configuredInterval = GetPrivateProfileIntW(
      L"watcher",
      L"monitor_interval_ms",
      kDefaultMonitorIntervalMs,
      g_state.configPath.c_str());
  g_state.monitorIntervalMs = ClampMonitorInterval(configuredInterval);

  wchar_t unitBuffer[32] = {};
  const DWORD unitCharsRead = GetPrivateProfileStringW(
      L"watcher",
      L"monitor_interval_unit",
      L"",
      unitBuffer,
      static_cast<DWORD>(ARRAYSIZE(unitBuffer)),
      g_state.configPath.c_str());

  if (unitCharsRead == 0) {
    g_state.monitorIntervalUseMinutes = IsWholeMinutesInterval(g_state.monitorIntervalMs);
    return;
  }

  if (lstrcmpiW(unitBuffer, L"minutes") == 0) {
    g_state.monitorIntervalUseMinutes = true;
    return;
  }
  if (lstrcmpiW(unitBuffer, L"seconds") == 0) {
    g_state.monitorIntervalUseMinutes = false;
    return;
  }

  g_state.monitorIntervalUseMinutes = IsWholeMinutesInterval(g_state.monitorIntervalMs);
  SaveMonitorIntervalUnitToConfig(g_state.monitorIntervalUseMinutes);
}

void SaveDoubleClickActionToConfig(DoubleClickAction action) {
  WritePrivateProfileStringW(
      L"watcher",
      L"double_click_action",
      DoubleClickActionConfigValue(action),
      g_state.configPath.c_str());
}

void LoadDoubleClickActionFromConfig() {
  wchar_t actionBuffer[64] = {};
  const DWORD charsRead = GetPrivateProfileStringW(
      L"watcher",
      L"double_click_action",
      L"",
      actionBuffer,
      static_cast<DWORD>(ARRAYSIZE(actionBuffer)),
      g_state.configPath.c_str());

  DoubleClickAction parsedAction = DoubleClickAction::kOpenLogFile;
  if (charsRead == 0 || !TryParseDoubleClickAction(actionBuffer, &parsedAction)) {
    g_state.doubleClickAction = DoubleClickAction::kOpenLogFile;
    SaveDoubleClickActionToConfig(g_state.doubleClickAction);
    return;
  }

  g_state.doubleClickAction = parsedAction;
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
  LoadDoubleClickActionFromConfig();
  LoadAcknowledgePopupDurationFromConfig();
  LoadAcknowledgedOffsetFromConfig();

  DebugLog(
      L"Config loaded. logPath=" + g_state.logPath +
      L", monitorIntervalMs=" + std::to_wstring(g_state.monitorIntervalMs) +
      L", useMinutes=" + std::to_wstring(g_state.monitorIntervalUseMinutes ? 1 : 0) +
      L", doubleClickAction=" + DoubleClickActionLabel(g_state.doubleClickAction) +
      L", acknowledgedOffset=" + std::to_wstring(g_state.acknowledgedOffset));
}

bool TryQueryIgnoreListState(bool* outExists, std::filesystem::file_time_type* outLastWriteTime) {
  if (outExists) {
    *outExists = false;
  }
  if (outLastWriteTime) {
    *outLastWriteTime = {};
  }

  std::error_code error;
  const std::filesystem::path ignorePath(g_state.ignorePath);
  const bool exists = std::filesystem::exists(ignorePath, error);
  if (error) {
    return false;
  }

  if (outExists) {
    *outExists = exists;
  }
  if (outLastWriteTime && exists) {
    *outLastWriteTime = std::filesystem::last_write_time(ignorePath, error);
    if (error) {
      return false;
    }
  }

  return true;
}

void UpdateIgnoreListStateCache(bool exists, const std::filesystem::file_time_type& lastWriteTime) {
  g_state.ignoreListStateKnown = true;
  g_state.ignoreFileExists = exists;
  g_state.ignoreFileLastWriteTime = lastWriteTime;
}

void RefreshIgnoreListStateCache() {
  bool exists = false;
  std::filesystem::file_time_type lastWriteTime = {};
  if (TryQueryIgnoreListState(&exists, &lastWriteTime)) {
    UpdateIgnoreListStateCache(exists, lastWriteTime);
  } else {
    g_state.ignoreListStateKnown = false;
  }
}

void LoadIgnoreList() {
  g_state.ignoredRules.clear();

  std::ifstream inputFile(std::filesystem::path(g_state.ignorePath), std::ios::binary);
  if (!inputFile.is_open()) {
    DebugLog(L"Ignore list file not found. path=" + g_state.ignorePath);
    return;
  }

  std::string line;
  while (std::getline(inputFile, line)) {
    if (!line.empty() && line.back() == '\r') {
      line.pop_back();
    }
    IgnoreRule rule = {};
    if (TryBuildIgnoreRule(line, &rule)) {
      g_state.ignoredRules.push_back(std::move(rule));
    }
  }

  DebugLog(L"Ignore list loaded. count=" + std::to_wstring(g_state.ignoredRules.size()));
}

void SaveIgnoreList() {
  std::ofstream outputFile(std::filesystem::path(g_state.ignorePath), std::ios::binary | std::ios::trunc);
  if (!outputFile.is_open()) {
    DebugLog(L"Failed to open ignore list for writing. path=" + g_state.ignorePath);
    return;
  }

  for (const IgnoreRule& rule : g_state.ignoredRules) {
    const std::string serializedRule = SerializeIgnoreRule(rule);
    outputFile.write(serializedRule.data(), static_cast<std::streamsize>(serializedRule.size()));
    outputFile.put('\n');
  }

  outputFile.flush();
  RefreshIgnoreListStateCache();
  DebugLog(L"Ignore list saved. count=" + std::to_wstring(g_state.ignoredRules.size()));
}

bool ReloadIgnoreListIfChanged(bool forceReload) {
  bool exists = false;
  std::filesystem::file_time_type lastWriteTime = {};
  if (!TryQueryIgnoreListState(&exists, &lastWriteTime)) {
    DebugLog(L"Failed to query Ignore.txt metadata.");
    return false;
  }

  if (!forceReload && g_state.ignoreListStateKnown &&
      g_state.ignoreFileExists == exists &&
      (!exists || g_state.ignoreFileLastWriteTime == lastWriteTime)) {
    return false;
  }

  LoadIgnoreList();
  UpdateIgnoreListStateCache(exists, lastWriteTime);
  return true;
}

bool EnsureIgnoreListFileExists() {
  std::error_code error;
  const std::filesystem::path ignorePath(g_state.ignorePath);
  if (std::filesystem::exists(ignorePath, error)) {
    RefreshIgnoreListStateCache();
    return true;
  }
  if (error) {
    return false;
  }

  std::ofstream outputFile(ignorePath, std::ios::binary | std::ios::app);
  if (!outputFile.is_open()) {
    return false;
  }

  outputFile.close();
  RefreshIgnoreListStateCache();
  return true;
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

  const bool alreadyRunningValue = alreadyRunning ? *alreadyRunning : false;
  DebugLog(L"AcquireSingleInstanceLock result alreadyRunning=" + std::to_wstring(alreadyRunningValue ? 1 : 0));
  return true;
}

void ReleaseSingleInstanceLock() {
  if (g_state.singleInstanceMutex) {
    CloseHandle(g_state.singleInstanceMutex);
    g_state.singleInstanceMutex = nullptr;
    DebugLog(L"Single-instance lock released.");
  }
}

void RefreshTrayIconVisualState() {
  std::wstring tip = L"Backrest Watcher: ";
  HICON icon = g_state.normalIcon;
  switch (g_state.alertSeverity) {
    case AlertSeverity::kWarning:
      tip += L"WARNING";
      icon = g_state.warningIcon;
      break;
    case AlertSeverity::kError:
      tip += L"ERROR";
      icon = g_state.blinkShowAlertIcon ? g_state.errorIcon : g_state.normalIcon;
      break;
    case AlertSeverity::kNone:
    default:
      tip += L"OK";
      break;
  }

  StringCchCopyW(g_state.trayIcon.szTip, ARRAYSIZE(g_state.trayIcon.szTip), tip.c_str());
  g_state.trayIcon.hIcon = icon;
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

bool TryGetLogFileSize(ULONGLONG* outSize) {
  if (!outSize) {
    return false;
  }

  HANDLE file = CreateFileW(
      g_state.logPath.c_str(),
      GENERIC_READ,
      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
      nullptr,
      OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL,
      nullptr);
  if (file == INVALID_HANDLE_VALUE) {
    return false;
  }

  LARGE_INTEGER fileSize = {};
  const bool ok = GetFileSizeEx(file, &fileSize) && fileSize.QuadPart >= 0;
  CloseHandle(file);
  if (!ok) {
    return false;
  }

  *outSize = static_cast<ULONGLONG>(fileSize.QuadPart);
  return true;
}

void UpdateTrayIcon() {
  RefreshTrayIconVisualState();
  g_state.trayIcon.uFlags = NIF_ICON | NIF_TIP;
  if (!Shell_NotifyIconW(NIM_MODIFY, &g_state.trayIcon)) {
    AddTrayIcon();
  }
}

void RefreshAlertStateFromEntries() {
  AlertSeverity highestSeverity = AlertSeverity::kNone;
  for (const AlertEntry& entry : g_state.activeAlertEntries) {
    if (!entry.isIgnored) {
      highestSeverity = MaxAlertSeverity(highestSeverity, entry.severity);
    }
  }
  g_state.alertSeverity = highestSeverity;
  g_state.blinkShowAlertIcon = true;
  UpdateTrayIcon();
}

bool AddIgnoredAlertRule(std::string_view rawLine) {
  IgnoreRule rule = {};
  if (!TryBuildIgnoreRule(rawLine, &rule)) {
    DebugLog(L"AddIgnoredAlertRule skipped.");
    return false;
  }

  const bool alreadyExists = std::any_of(
      g_state.ignoredRules.begin(),
      g_state.ignoredRules.end(),
      [&rule](const IgnoreRule& existingRule) {
        return IsSameIgnoreRule(existingRule, rule);
      });
  if (alreadyExists) {
    DebugLog(L"AddIgnoredAlertRule skipped because it already exists.");
    return false;
  }

  g_state.ignoredRules.push_back(std::move(rule));
  SaveIgnoreList();
  DebugLog(L"Ignored rule added.");
  return true;
}

bool RemoveIgnoredAlertRuleAt(size_t index) {
  if (index >= g_state.ignoredRules.size()) {
    return false;
  }

  using IgnoreVector = std::vector<IgnoreRule>;
  g_state.ignoredRules.erase(
      g_state.ignoredRules.begin() + static_cast<IgnoreVector::difference_type>(index));
  SaveIgnoreList();
  DebugLog(L"Ignored rule removed. newCount=" + std::to_wstring(g_state.ignoredRules.size()));
  return true;
}

LRESULT CALLBACK AcknowledgePopupWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  DebugLogWindowMessage(L"AcknowledgePopupWindow", message, wParam, lParam);
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

  popupX = (std::max)(
      static_cast<int>(workArea.left),
      (std::min)(popupX, static_cast<int>(workArea.right) - kAcknowledgePopupWidth));
  popupY = (std::max)(
      static_cast<int>(workArea.top),
      (std::min)(popupY, static_cast<int>(workArea.bottom) - kAcknowledgePopupHeight));

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
  DebugLog(L"Acknowledge popup shown.");
}

void ResetWatcherAndRescan() {
  DebugLog(L"ResetWatcherAndRescan started.");
  g_state.lastOffset = 0;
  g_state.lastLineNumber = 0;
  g_state.activeAlertEntries.clear();
  g_state.alertSeverity = AlertSeverity::kNone;
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
      const ULONGLONG startingLineNumber = StartingLineNumberForOffset(file, g_state.acknowledgedOffset);
      g_state.alertSeverity = ScanFileRangeForAlertEntries(
          file,
          g_state.acknowledgedOffset,
          currentSize,
          startingLineNumber,
          &g_state.lastLineNumber,
          &g_state.activeAlertEntries);
      g_state.lastOffset = currentSize;
    }
    CloseHandle(file);
  }

  UpdateTrayIcon();
  RefreshAlertManagerWindowContent();
  DebugLog(
      L"ResetWatcherAndRescan finished. severity=" + std::wstring(AlertSeverityLabel(g_state.alertSeverity)) +
      L", entries=" + std::to_wstring(g_state.activeAlertEntries.size()) +
      L", lastOffset=" + std::to_wstring(g_state.lastOffset) +
      L", lastLineNumber=" + std::to_wstring(g_state.lastLineNumber));
}

void MonitorLogFileOnce() {
  DebugLog(L"MonitorLogFileOnce tick started.");
  if (ReloadIgnoreListIfChanged(false)) {
    DebugLog(L"Ignore list changed on disk. Rescanning log.");
    ResetWatcherAndRescan();
    return;
  }

  HANDLE file = CreateFileW(
      g_state.logPath.c_str(),
      GENERIC_READ,
      FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
      nullptr,
      OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL,
      nullptr);

  if (file == INVALID_HANDLE_VALUE) {
    const bool hadActiveAlerts = !g_state.activeAlertEntries.empty();
    g_state.activeAlertEntries.clear();
    if (HasAlert(g_state.alertSeverity) || hadActiveAlerts) {
      g_state.alertSeverity = AlertSeverity::kNone;
      g_state.blinkShowAlertIcon = true;
      UpdateTrayIcon();
      RefreshAlertManagerWindowContent();
    }
    g_state.lastOffset = 0;
    g_state.lastLineNumber = 0;
    DebugLog(L"MonitorLogFileOnce: log file unavailable.");
    return;
  }

  LARGE_INTEGER fileSize = {};
  if (!GetFileSizeEx(file, &fileSize)) {
    CloseHandle(file);
    return;
  }

  const ULONGLONG newSize = static_cast<ULONGLONG>(fileSize.QuadPart);
  bool needIconRefresh = false;
  bool needAlertWindowRefresh = false;

  if (newSize < g_state.lastOffset) {
    if (g_state.acknowledgedOffset > newSize) {
      g_state.acknowledgedOffset = 0;
      SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
    }
    g_state.lastOffset = 0;
    g_state.lastLineNumber = 0;
    g_state.activeAlertEntries.clear();
    g_state.alertSeverity = AlertSeverity::kNone;
    g_state.blinkShowAlertIcon = true;
    needIconRefresh = true;
    needAlertWindowRefresh = true;
  }

  if (newSize > g_state.lastOffset) {
    std::vector<AlertEntry> newEntries;
    ULONGLONG endingLineNumber = g_state.lastLineNumber;
    const AlertSeverity newSeverity = ScanFileRangeForAlertEntries(
        file,
        g_state.lastOffset,
        newSize,
        g_state.lastLineNumber,
        &endingLineNumber,
        &newEntries);
    const AlertSeverity updatedSeverity = MaxAlertSeverity(g_state.alertSeverity, newSeverity);
    if (updatedSeverity != g_state.alertSeverity) {
      g_state.alertSeverity = updatedSeverity;
      if (HasAlert(g_state.alertSeverity)) {
        g_state.blinkShowAlertIcon = true;
      }
      needIconRefresh = true;
    }
    if (!newEntries.empty()) {
      g_state.activeAlertEntries.insert(
          g_state.activeAlertEntries.end(),
          std::make_move_iterator(newEntries.begin()),
          std::make_move_iterator(newEntries.end()));
        needAlertWindowRefresh = true;
      }
      g_state.lastOffset = newSize;
      g_state.lastLineNumber = endingLineNumber;
    }

  if (needIconRefresh) {
    UpdateTrayIcon();
  }
  if (needAlertWindowRefresh) {
    RefreshAlertManagerWindowContent();
  }

  CloseHandle(file);
  DebugLog(
      L"MonitorLogFileOnce finished. severity=" + std::wstring(AlertSeverityLabel(g_state.alertSeverity)) +
      L", entries=" + std::to_wstring(g_state.activeAlertEntries.size()) +
      L", lastOffset=" + std::to_wstring(g_state.lastOffset));
}

void OpenLogFolder() {
  DebugLog(L"OpenLogFolder requested. path=" + g_state.logPath);
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
  DebugLog(L"OpenLogFile requested. path=" + g_state.logPath);
  const HINSTANCE result = ShellExecuteW(nullptr, L"open", g_state.logPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(result) <= 32) {
    MessageBoxW(g_state.hwnd, L"Cannot open log file.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
  }
}

void OpenIgnoreListFile() {
  DebugLog(L"OpenIgnoreListFile requested. path=" + g_state.ignorePath);
  if (!EnsureIgnoreListFileExists()) {
    MessageBoxW(g_state.hwnd, L"Cannot create Ignore.txt.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    return;
  }

  const HINSTANCE result = ShellExecuteW(nullptr, L"open", g_state.ignorePath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if (reinterpret_cast<INT_PTR>(result) <= 32) {
    MessageBoxW(g_state.hwnd, L"Cannot open Ignore.txt.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
  }
}

void AcknowledgeAlert() {
  DebugLog(L"AcknowledgeAlert requested.");
  ULONGLONG currentLogSize = 0;
  if (TryGetLogFileSize(&currentLogSize)) {
    g_state.lastOffset = currentLogSize;
    g_state.acknowledgedOffset = currentLogSize;
    HANDLE file = CreateFileW(
        g_state.logPath.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file != INVALID_HANDLE_VALUE) {
      g_state.lastLineNumber = CountLogicalLinesUpToOffset(file, currentLogSize);
      CloseHandle(file);
    }
  } else {
    g_state.acknowledgedOffset = g_state.lastOffset;
    if (g_state.lastOffset == 0) {
      g_state.lastLineNumber = 0;
    }
  }
  SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
  g_state.activeAlertEntries.clear();
  g_state.alertSeverity = AlertSeverity::kNone;
  g_state.blinkShowAlertIcon = true;
  UpdateTrayIcon();
  RefreshAlertManagerWindowContent();
  ShowAcknowledgeNotification();
  DebugLog(L"AcknowledgeAlert finished. acknowledgedOffset=" + std::to_wstring(g_state.acknowledgedOffset));
}

void ApplyMonitorInterval() {
  if (!g_state.hwnd) {
    return;
  }
  KillTimer(g_state.hwnd, kMonitorTimerId);
  if (SetTimer(g_state.hwnd, kMonitorTimerId, g_state.monitorIntervalMs, nullptr) == 0) {
    MessageBoxW(g_state.hwnd, L"Cannot update log monitoring timer.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
  }
  DebugLog(L"ApplyMonitorInterval set timer to " + std::to_wstring(g_state.monitorIntervalMs) + L" ms.");
}

void SetMonitorInterval(UINT intervalMs, bool useMinutes) {
  g_state.monitorIntervalMs = ClampMonitorInterval(intervalMs);
  g_state.monitorIntervalUseMinutes = useMinutes;
  SaveMonitorIntervalToConfig(g_state.monitorIntervalMs);
  SaveMonitorIntervalUnitToConfig(g_state.monitorIntervalUseMinutes);
  ApplyMonitorInterval();
  DebugLog(
      L"SetMonitorInterval applied. intervalMs=" + std::to_wstring(g_state.monitorIntervalMs) +
      L", useMinutes=" + std::to_wstring(g_state.monitorIntervalUseMinutes ? 1 : 0));
}

LRESULT CALLBACK IntervalInputWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  DebugLogWindowMessage(L"IntervalInputWindow", message, wParam, lParam);
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

      const bool useMinutes = state->initialUseMinutes;
      SendMessageW(state->secondsCheckbox, BM_SETCHECK, useMinutes ? BST_UNCHECKED : BST_CHECKED, 0);
      SendMessageW(state->minutesCheckbox, BM_SETCHECK, useMinutes ? BST_CHECKED : BST_UNCHECKED, 0);

      wchar_t initialBuffer[32] = {};
      if (useMinutes) {
        StringCchPrintfW(initialBuffer, ARRAYSIZE(initialBuffer), L"%.2f", static_cast<double>(state->initialValueMs) / 60000.0);
      } else {
        StringCchPrintfW(initialBuffer, ARRAYSIZE(initialBuffer), L"%.2f", static_cast<double>(state->initialValueMs) / 1000.0);
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
        state->resultUseMinutes = useMinutes;
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

bool PromptMonitorIntervalMs(HWND parent, UINT currentValueMs, bool currentUseMinutes, UINT* outValueMs, bool* outUseMinutes) {
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
  state.initialUseMinutes = currentUseMinutes;

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

  const int dialogWidth = 300;
  const int dialogHeight = 170;
  CenterWindowOnCurrentScreen(dialog, dialogWidth, dialogHeight, SWP_NOZORDER | SWP_SHOWWINDOW);

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
  if (outUseMinutes) {
    *outUseMinutes = state.resultUseMinutes;
  }
  return true;
}

void PromptAndSetMonitorInterval() {
  UINT intervalMs = g_state.monitorIntervalMs;
  bool useMinutes = g_state.monitorIntervalUseMinutes;
  if (!PromptMonitorIntervalMs(
          g_state.hwnd,
          g_state.monitorIntervalMs,
          g_state.monitorIntervalUseMinutes,
          &intervalMs,
          &useMinutes)) {
    DebugLog(L"PromptAndSetMonitorInterval canceled.");
    return;
  }
  SetMonitorInterval(intervalMs, useMinutes);
  DebugLog(L"PromptAndSetMonitorInterval accepted.");
}

LRESULT CALLBACK DoubleClickActionWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  DebugLogWindowMessage(L"DoubleClickActionWindow", message, wParam, lParam);
  auto* state = reinterpret_cast<DoubleClickActionDialogState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

  switch (message) {
    case WM_CREATE: {
      auto* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
      state = reinterpret_cast<DoubleClickActionDialogState*>(createStruct->lpCreateParams);
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));

      CreateWindowExW(
          0,
          L"STATIC",
          L"Choose what double-left-click should do:",
          WS_CHILD | WS_VISIBLE,
          12,
          14,
          324,
          20,
          hwnd,
          nullptr,
          nullptr,
          nullptr);

      state->comboBox = CreateWindowExW(
          0,
          L"COMBOBOX",
          L"",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
          12,
          40,
          324,
          240,
          hwnd,
          MenuHandleFromId(kControlDoubleClickActionCombo),
          nullptr,
          nullptr);

      int selectedIndex = 0;
      for (int i = 0; i < static_cast<int>(ARRAYSIZE(kDoubleClickActionOptions)); ++i) {
        const DoubleClickAction action = kDoubleClickActionOptions[i];
        const int comboIndex = static_cast<int>(SendMessageW(
            state->comboBox,
            CB_ADDSTRING,
            0,
            reinterpret_cast<LPARAM>(DoubleClickActionLabel(action))));
        if (comboIndex >= 0) {
          SendMessageW(
              state->comboBox,
              CB_SETITEMDATA,
              static_cast<WPARAM>(comboIndex),
              static_cast<LPARAM>(action));
          if (action == state->initialAction) {
            selectedIndex = comboIndex;
          }
        }
      }
      SendMessageW(state->comboBox, CB_SETCURSEL, static_cast<WPARAM>(selectedIndex), 0);

      CreateWindowExW(
          0,
          L"BUTTON",
          L"OK",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
          178,
          86,
          74,
          26,
          hwnd,
          reinterpret_cast<HMENU>(IDOK),
          nullptr,
          nullptr);

      CreateWindowExW(
          0,
          L"BUTTON",
          L"Cancel",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
          262,
          86,
          74,
          26,
          hwnd,
          reinterpret_cast<HMENU>(IDCANCEL),
          nullptr,
          nullptr);

      SetFocus(state->comboBox);
      return 0;
    }

    case WM_COMMAND:
      if (LOWORD(wParam) == IDOK) {
        const int selectedIndex = static_cast<int>(SendMessageW(state->comboBox, CB_GETCURSEL, 0, 0));
        if (selectedIndex == CB_ERR) {
          MessageBoxW(hwnd, L"Choose an action first.", L"Backrest Watcher", MB_ICONWARNING | MB_OK);
          return 0;
        }

        const LRESULT itemData = SendMessageW(state->comboBox, CB_GETITEMDATA, selectedIndex, 0);
        if (itemData == CB_ERR) {
          return 0;
        }

        state->resultAction = static_cast<DoubleClickAction>(itemData);
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

    default:
      return DefWindowProcW(hwnd, message, wParam, lParam);
  }
}

bool PromptDoubleClickAction(HWND parent, DoubleClickAction currentAction, DoubleClickAction* outAction) {
  static bool registered = false;
  if (!registered) {
    WNDCLASSW cls = {};
    cls.lpfnWndProc = DoubleClickActionWindowProc;
    cls.hInstance = GetModuleHandleW(nullptr);
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    cls.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    cls.lpszClassName = kDoubleClickActionDialogClassName;
    RegisterClassW(&cls);
    registered = true;
  }

  DoubleClickActionDialogState state = {};
  state.initialAction = currentAction;
  state.resultAction = currentAction;

  HWND dialog = CreateWindowExW(
      WS_EX_DLGMODALFRAME,
      kDoubleClickActionDialogClassName,
      L"Double-Click Action",
      WS_CAPTION | WS_POPUP | WS_SYSMENU,
      CW_USEDEFAULT,
      CW_USEDEFAULT,
      360,
      160,
      parent,
      nullptr,
      GetModuleHandleW(nullptr),
      &state);

  if (!dialog) {
    return false;
  }

  const int dialogWidth = 360;
  const int dialogHeight = 160;
  CenterWindowOnCurrentScreen(dialog, dialogWidth, dialogHeight, SWP_NOZORDER | SWP_SHOWWINDOW);

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

  if (outAction) {
    *outAction = state.resultAction;
  }
  return true;
}

void PromptAndSetDoubleClickAction() {
  DoubleClickAction selectedAction = g_state.doubleClickAction;
  if (!PromptDoubleClickAction(g_state.hwnd, g_state.doubleClickAction, &selectedAction)) {
    DebugLog(L"PromptAndSetDoubleClickAction canceled.");
    return;
  }

  g_state.doubleClickAction = selectedAction;
  SaveDoubleClickActionToConfig(g_state.doubleClickAction);
  DebugLog(L"PromptAndSetDoubleClickAction accepted. action=" + std::wstring(DoubleClickActionLabel(g_state.doubleClickAction)));
}

size_t SelectedListItemDataIndex(HWND listBoxHwnd, size_t maxValidSize) {
  if (!listBoxHwnd) {
    return maxValidSize;
  }

  const int selectedIndex = static_cast<int>(SendMessageW(listBoxHwnd, LB_GETCURSEL, 0, 0));
  if (selectedIndex == LB_ERR) {
    return maxValidSize;
  }

  const LRESULT itemData = SendMessageW(listBoxHwnd, LB_GETITEMDATA, selectedIndex, 0);
  if (itemData == LB_ERR) {
    return maxValidSize;
  }

  const size_t resolvedIndex = static_cast<size_t>(itemData);
  if (resolvedIndex >= maxValidSize) {
    return maxValidSize;
  }
  return resolvedIndex;
}

void UpdateActiveAlertsSelectionDetails() {
  if (!g_state.activeAlertsDetailsHwnd || !g_state.ignoreSelectedButtonHwnd) {
    return;
  }

  if (g_state.activeAlertEntries.empty()) {
    SetWindowTextW(
        g_state.activeAlertsDetailsHwnd,
        L"No warn/error messages are waiting right now.");
    EnableWindow(g_state.ignoreSelectedButtonHwnd, FALSE);
    return;
  }

  const size_t selectedIndex = SelectedListItemDataIndex(
      g_state.activeAlertsListHwnd,
      g_state.activeAlertEntries.size());
  if (selectedIndex >= g_state.activeAlertEntries.size()) {
    SetWindowTextW(
        g_state.activeAlertsDetailsHwnd,
        L"Select a message above to view details. Double-click a message to jump to that line in backrest.log.");
    EnableWindow(g_state.ignoreSelectedButtonHwnd, FALSE);
    return;
  }

  const AlertEntry& entry = g_state.activeAlertEntries[selectedIndex];
  SetWindowTextW(g_state.activeAlertsDetailsHwnd, entry.detailText.c_str());
  EnableWindow(g_state.ignoreSelectedButtonHwnd, entry.isIgnored ? FALSE : TRUE);
}

void UpdateIgnoredAlertsSelectionDetails() {
  if (!g_state.ignoredAlertsDetailsHwnd || !g_state.refreshIgnoredButtonHwnd) {
    return;
  }

  if (g_state.ignoredRules.empty()) {
    SetWindowTextW(
        g_state.ignoredAlertsDetailsHwnd,
        L"Ignore list is empty. Add rules in Ignore.txt, then click Refresh list.");
    EnableWindow(g_state.refreshIgnoredButtonHwnd, TRUE);
    return;
  }

  const size_t selectedIndex = SelectedListItemDataIndex(
      g_state.ignoredAlertsListHwnd,
      g_state.ignoredRules.size());
  if (selectedIndex >= g_state.ignoredRules.size()) {
    SetWindowTextW(
        g_state.ignoredAlertsDetailsHwnd,
        L"Select an ignore rule above to view its details. Click Refresh list after editing Ignore.txt.");
    EnableWindow(g_state.refreshIgnoredButtonHwnd, TRUE);
    return;
  }

  const std::wstring details = IgnoreRuleDetailsText(g_state.ignoredRules[selectedIndex]);
  SetWindowTextW(g_state.ignoredAlertsDetailsHwnd, details.c_str());
  EnableWindow(g_state.refreshIgnoredButtonHwnd, TRUE);
}

void UpdateAlertManagerVisibleTab() {
  if (!g_state.alertManagerTabHwnd) {
    return;
  }

  const int currentTab = TabCtrl_GetCurSel(g_state.alertManagerTabHwnd);
  const BOOL showMessages = (currentTab != static_cast<int>(kAlertManagerTabIgnored)) ? TRUE : FALSE;

  ShowWindow(g_state.activeAlertsListHwnd, showMessages ? SW_SHOW : SW_HIDE);
  ShowWindow(g_state.activeAlertsDetailsHwnd, showMessages ? SW_SHOW : SW_HIDE);
  ShowWindow(g_state.ignoreSelectedButtonHwnd, showMessages ? SW_SHOW : SW_HIDE);
  ShowWindow(g_state.ignoredAlertsListHwnd, showMessages ? SW_HIDE : SW_SHOW);
  ShowWindow(g_state.ignoredAlertsDetailsHwnd, showMessages ? SW_HIDE : SW_SHOW);
  ShowWindow(g_state.refreshIgnoredButtonHwnd, showMessages ? SW_HIDE : SW_SHOW);
  ShowWindow(g_state.openIgnoreFileButtonHwnd, showMessages ? SW_HIDE : SW_SHOW);
  ShowWindow(g_state.ignoreUsageHintHwnd, showMessages ? SW_HIDE : SW_SHOW);

  if (showMessages) {
    UpdateActiveAlertsSelectionDetails();
  } else {
    UpdateIgnoredAlertsSelectionDetails();
  }
}

void LayoutAlertManagerControls() {
  if (!g_state.alertManagerHwnd || !g_state.alertManagerTabHwnd) {
    return;
  }

  RECT clientRect = {};
  GetClientRect(g_state.alertManagerHwnd, &clientRect);

  const int tabX = kAlertManagerPaddingPx;
  const int tabY = kAlertManagerPaddingPx;
  const int tabWidth = (std::max)(100, static_cast<int>(clientRect.right) - (kAlertManagerPaddingPx * 2));
  const int tabHeight = (std::max)(100, static_cast<int>(clientRect.bottom) - (kAlertManagerPaddingPx * 2));
  SetWindowPos(
      g_state.alertManagerTabHwnd,
      nullptr,
      tabX,
      tabY,
      tabWidth,
      tabHeight,
      SWP_NOZORDER);

  RECT tabContentRect = {tabX, tabY, tabX + tabWidth, tabY + tabHeight};
  TabCtrl_AdjustRect(g_state.alertManagerTabHwnd, FALSE, &tabContentRect);

  const int contentWidth = (std::max)(100, static_cast<int>(tabContentRect.right - tabContentRect.left));
  const int contentHeight = (std::max)(100, static_cast<int>(tabContentRect.bottom - tabContentRect.top));
  const int listWidth = (std::max)(240, contentWidth - kAlertManagerButtonWidth - kAlertManagerPaddingPx);
  const int listHeight = (std::max)(120, contentHeight - kAlertManagerDetailsHeight - kAlertManagerPaddingPx);
  const int detailsHeight = (std::max)(90, contentHeight - listHeight - kAlertManagerPaddingPx);
  const int buttonX = tabContentRect.left + listWidth + kAlertManagerPaddingPx;
  const int buttonY = tabContentRect.top;
  const int secondButtonY = buttonY + kAlertManagerButtonHeight + kAlertManagerButtonSpacingPx;
  const int usageHintY = secondButtonY + kAlertManagerButtonHeight + kAlertManagerButtonSpacingPx;
  const int usageHintHeight = (std::max)(
      0,
      listHeight - ((kAlertManagerButtonHeight * 2) + (kAlertManagerButtonSpacingPx * 2)));
  const int detailsY = tabContentRect.top + listHeight + kAlertManagerPaddingPx;

  SetWindowPos(
      g_state.activeAlertsListHwnd,
      nullptr,
      tabContentRect.left,
      tabContentRect.top,
      listWidth,
      listHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.ignoredAlertsListHwnd,
      nullptr,
      tabContentRect.left,
      tabContentRect.top,
      listWidth,
      listHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.ignoreSelectedButtonHwnd,
      nullptr,
      buttonX,
      buttonY,
      kAlertManagerButtonWidth,
      kAlertManagerButtonHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.refreshIgnoredButtonHwnd,
      nullptr,
      buttonX,
      buttonY,
      kAlertManagerButtonWidth,
      kAlertManagerButtonHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.openIgnoreFileButtonHwnd,
      nullptr,
      buttonX,
      secondButtonY,
      kAlertManagerButtonWidth,
      kAlertManagerButtonHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.ignoreUsageHintHwnd,
      nullptr,
      buttonX,
      usageHintY,
      kAlertManagerButtonWidth,
      usageHintHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.activeAlertsDetailsHwnd,
      nullptr,
      tabContentRect.left,
      detailsY,
      contentWidth,
      detailsHeight,
      SWP_NOZORDER);
  SetWindowPos(
      g_state.ignoredAlertsDetailsHwnd,
      nullptr,
      tabContentRect.left,
      detailsY,
      contentWidth,
      detailsHeight,
      SWP_NOZORDER);
}

void RefreshAlertManagerWindowContent() {
  if (!g_state.alertManagerHwnd || !IsWindow(g_state.alertManagerHwnd)) {
    return;
  }

  if (g_state.activeAlertsListHwnd) {
    SendMessageW(g_state.activeAlertsListHwnd, LB_RESETCONTENT, 0, 0);
    for (int sourceIndex = static_cast<int>(g_state.activeAlertEntries.size()) - 1; sourceIndex >= 0; --sourceIndex) {
      const AlertEntry& entry = g_state.activeAlertEntries[sourceIndex];
      const int listIndex = static_cast<int>(SendMessageW(
          g_state.activeAlertsListHwnd,
          LB_ADDSTRING,
          0,
          reinterpret_cast<LPARAM>(entry.listText.c_str())));
      if (listIndex >= 0) {
        SendMessageW(g_state.activeAlertsListHwnd, LB_SETITEMDATA, static_cast<WPARAM>(listIndex), sourceIndex);
      }
    }
    UpdateListBoxHorizontalExtent(g_state.activeAlertsListHwnd);
    if (!g_state.activeAlertEntries.empty()) {
      SendMessageW(g_state.activeAlertsListHwnd, LB_SETCURSEL, 0, 0);
    }
  }

  if (g_state.ignoredAlertsListHwnd) {
    SendMessageW(g_state.ignoredAlertsListHwnd, LB_RESETCONTENT, 0, 0);
    for (size_t i = 0; i < g_state.ignoredRules.size(); ++i) {
      const std::wstring displayText = IgnoreRuleDisplayText(g_state.ignoredRules[i]);
      const int listIndex = static_cast<int>(SendMessageW(
          g_state.ignoredAlertsListHwnd,
          LB_ADDSTRING,
          0,
          reinterpret_cast<LPARAM>(displayText.c_str())));
      if (listIndex >= 0) {
        SendMessageW(g_state.ignoredAlertsListHwnd, LB_SETITEMDATA, static_cast<WPARAM>(listIndex), i);
      }
    }
    UpdateListBoxHorizontalExtent(g_state.ignoredAlertsListHwnd);
    if (!g_state.ignoredRules.empty()) {
      SendMessageW(g_state.ignoredAlertsListHwnd, LB_SETCURSEL, 0, 0);
    }
  }

  UpdateAlertManagerVisibleTab();
}

void OpenSelectedActiveAlertInLog() {
  const size_t selectedIndex = SelectedListItemDataIndex(
      g_state.activeAlertsListHwnd,
      g_state.activeAlertEntries.size());
  if (selectedIndex >= g_state.activeAlertEntries.size()) {
    return;
  }

  const AlertEntry& entry = g_state.activeAlertEntries[selectedIndex];
  DebugLog(
      L"OpenSelectedActiveAlertInLog requested. line=" + std::to_wstring(entry.lineNumber) +
      L", summary=" + entry.summaryText);
  OpenLogFileAtLine(entry.lineNumber);
}

void IgnoreSelectedActiveAlert() {
  const size_t selectedIndex = SelectedListItemDataIndex(
      g_state.activeAlertsListHwnd,
      g_state.activeAlertEntries.size());
  if (selectedIndex >= g_state.activeAlertEntries.size()) {
    MessageBoxW(
        g_state.alertManagerHwnd,
        L"Select a warn/error message to ignore first.",
        L"Backrest Watcher",
        MB_ICONWARNING | MB_OK);
    return;
  }

  const AlertEntry selectedEntry = g_state.activeAlertEntries[selectedIndex];
  if (selectedEntry.isIgnored) {
    MessageBoxW(
        g_state.alertManagerHwnd,
        L"This message already matches Ignore.txt.",
        L"Backrest Watcher",
        MB_ICONINFORMATION | MB_OK);
    return;
  }
  DebugLog(L"IgnoreSelectedActiveAlert requested for " + selectedEntry.summaryText);
  if (selectedEntry.severity == AlertSeverity::kError) {
    std::wstring confirmText =
        L"Ignore this error message?\r\n\r\nSummary:\r\n" +
        selectedEntry.summaryText;
    if (MessageBoxW(
            g_state.alertManagerHwnd,
            confirmText.c_str(),
            L"Confirm ignore",
            MB_ICONWARNING | MB_YESNO | MB_DEFBUTTON2) != IDYES) {
      return;
    }
  }

  AddIgnoredAlertRule(selectedEntry.ignoreRuleText);
  ResetWatcherAndRescan();
  DebugLog(L"IgnoreSelectedActiveAlert completed.");
}

void RefreshIgnoreListFromDisk() {
  DebugLog(L"RefreshIgnoreListFromDisk requested.");
  ReloadIgnoreListIfChanged(true);
  ResetWatcherAndRescan();
}

LRESULT CALLBACK AlertManagerWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
  DebugLogWindowMessage(L"AlertManagerWindow", message, wParam, lParam);
  switch (message) {
    case WM_CREATE: {
      g_state.alertManagerHwnd = hwnd;
      g_state.alertManagerTabHwnd = CreateWindowExW(
          0,
          WC_TABCONTROLW,
          L"",
          WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_TABSTOP,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlAlertTab),
          nullptr,
          nullptr);

      TCITEMW messagesItem = {};
      messagesItem.mask = TCIF_TEXT;
      messagesItem.pszText = const_cast<LPWSTR>(L"Messages");
      TabCtrl_InsertItem(g_state.alertManagerTabHwnd, static_cast<int>(kAlertManagerTabMessages), &messagesItem);

      TCITEMW ignoredItem = {};
      ignoredItem.mask = TCIF_TEXT;
      ignoredItem.pszText = const_cast<LPWSTR>(L"Ignore list");
      TabCtrl_InsertItem(g_state.alertManagerTabHwnd, static_cast<int>(kAlertManagerTabIgnored), &ignoredItem);

      g_state.activeAlertsListHwnd = CreateWindowExW(
          WS_EX_CLIENTEDGE,
          L"LISTBOX",
          L"",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | LBS_NOTIFY | WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlActiveAlertsList),
          nullptr,
          nullptr);
      g_state.activeAlertsDetailsHwnd = CreateWindowExW(
          WS_EX_CLIENTEDGE,
          L"EDIT",
          L"",
          WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlActiveAlertsDetails),
          nullptr,
          nullptr);
      g_state.ignoreSelectedButtonHwnd = CreateWindowExW(
          0,
          L"BUTTON",
          L"Ignore selected",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlIgnoreSelectedButton),
          nullptr,
          nullptr);

      g_state.ignoredAlertsListHwnd = CreateWindowExW(
          WS_EX_CLIENTEDGE,
          L"LISTBOX",
          L"",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP | LBS_NOTIFY | WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlIgnoredAlertsList),
          nullptr,
          nullptr);
      g_state.ignoredAlertsDetailsHwnd = CreateWindowExW(
          WS_EX_CLIENTEDGE,
          L"EDIT",
          L"",
          WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlIgnoredAlertsDetails),
          nullptr,
          nullptr);
      g_state.refreshIgnoredButtonHwnd = CreateWindowExW(
          0,
          L"BUTTON",
          L"Refresh list",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlRefreshIgnoredButton),
          nullptr,
          nullptr);
      g_state.openIgnoreFileButtonHwnd = CreateWindowExW(
          0,
          L"BUTTON",
          L"Open Ignore.txt",
          WS_CHILD | WS_VISIBLE | WS_TABSTOP,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlOpenIgnoreFileButton),
          nullptr,
          nullptr);
      g_state.ignoreUsageHintHwnd = CreateWindowExW(
          WS_EX_CLIENTEDGE,
          L"EDIT",
          kIgnoreUsageHintText,
          WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
          0,
          0,
          0,
          0,
          hwnd,
          MenuHandleFromId(kControlIgnoreUsageHint),
          nullptr,
          nullptr);

      LayoutAlertManagerControls();
      RefreshAlertManagerWindowContent();
      return 0;
    }

    case WM_SIZE:
      LayoutAlertManagerControls();
      return 0;

    case WM_NOTIFY:
      if (reinterpret_cast<NMHDR*>(lParam)->idFrom == kControlAlertTab &&
          reinterpret_cast<NMHDR*>(lParam)->code == TCN_SELCHANGE) {
        UpdateAlertManagerVisibleTab();
        return 0;
      }
      break;

    case WM_COMMAND:
      switch (LOWORD(wParam)) {
        case kControlActiveAlertsList:
          if (HIWORD(wParam) == LBN_SELCHANGE) {
            UpdateActiveAlertsSelectionDetails();
          } else if (HIWORD(wParam) == LBN_DBLCLK) {
            OpenSelectedActiveAlertInLog();
          }
          return 0;
        case kControlIgnoredAlertsList:
          if (HIWORD(wParam) == LBN_SELCHANGE) {
            UpdateIgnoredAlertsSelectionDetails();
          }
          return 0;
        case kControlIgnoreSelectedButton:
          IgnoreSelectedActiveAlert();
          return 0;
        case kControlRefreshIgnoredButton:
          RefreshIgnoreListFromDisk();
          return 0;
        case kControlOpenIgnoreFileButton:
          OpenIgnoreListFile();
          return 0;
        default:
          return DefWindowProcW(hwnd, message, wParam, lParam);
      }

    case WM_CLOSE:
      DestroyWindow(hwnd);
      return 0;

    case WM_DESTROY:
      g_state.alertManagerHwnd = nullptr;
      g_state.alertManagerTabHwnd = nullptr;
      g_state.activeAlertsListHwnd = nullptr;
      g_state.activeAlertsDetailsHwnd = nullptr;
      g_state.ignoreSelectedButtonHwnd = nullptr;
      g_state.ignoredAlertsListHwnd = nullptr;
      g_state.ignoredAlertsDetailsHwnd = nullptr;
      g_state.refreshIgnoredButtonHwnd = nullptr;
      g_state.openIgnoreFileButtonHwnd = nullptr;
      g_state.ignoreUsageHintHwnd = nullptr;
      return 0;

    default:
      break;
  }

  return DefWindowProcW(hwnd, message, wParam, lParam);
}

bool EnsureAlertManagerWindowClassRegistered() {
  static bool registered = false;
  if (registered) {
    return true;
  }

  WNDCLASSW windowClass = {};
  windowClass.lpfnWndProc = AlertManagerWindowProc;
  windowClass.hInstance = GetModuleHandleW(nullptr);
  windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
  windowClass.lpszClassName = kAlertManagerWindowClassName;

  if (!RegisterClassW(&windowClass) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
    return false;
  }

  registered = true;
  return true;
}

void ShowAlertManagerWindow() {
  if (ReloadIgnoreListIfChanged(false)) {
    DebugLog(L"ShowAlertManagerWindow noticed Ignore.txt changed. Rescanning log.");
    ResetWatcherAndRescan();
  }

  if (g_state.alertManagerHwnd && IsWindow(g_state.alertManagerHwnd)) {
    ShowWindow(g_state.alertManagerHwnd, SW_SHOWNORMAL);
    SetForegroundWindow(g_state.alertManagerHwnd);
    RefreshAlertManagerWindowContent();
    DebugLog(L"ShowAlertManagerWindow reused existing window.");
    return;
  }

  if (!EnsureAlertManagerWindowClassRegistered()) {
    MessageBoxW(g_state.hwnd, L"Cannot create alert manager window.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    return;
  }

  const POINT alertWindowPosition = CenteredWindowPosition(kAlertManagerWindowWidth, kAlertManagerWindowHeight);
  HWND alertWindow = CreateWindowExW(
      0,
      kAlertManagerWindowClassName,
      L"Backrest Watcher Alerts",
      WS_OVERLAPPEDWINDOW,
      alertWindowPosition.x,
      alertWindowPosition.y,
      kAlertManagerWindowWidth,
      kAlertManagerWindowHeight,
      g_state.hwnd,
      nullptr,
      GetModuleHandleW(nullptr),
      nullptr);
  if (!alertWindow) {
    MessageBoxW(g_state.hwnd, L"Cannot create alert manager window.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    return;
  }

  g_state.alertManagerHwnd = alertWindow;
  ShowWindow(alertWindow, SW_SHOWNORMAL);
  UpdateWindow(alertWindow);
  DebugLog(L"ShowAlertManagerWindow created a new window.");
}

void ApplySelectedLogPath(const std::wstring& selectedPath) {
  g_state.logPath = selectedPath;
  SaveLogPathToConfig(g_state.logPath);
  g_state.acknowledgedOffset = 0;
  SaveAcknowledgedOffsetToConfig(g_state.acknowledgedOffset);
  ResetWatcherAndRescan();
}

void ChooseLogPath() {
  DebugLog(L"ChooseLogPath opened.");
  if (g_state.debugMode) {
    const std::wstring automatedPath =
        ReadEnvironmentVariableValue(L"BACKREST_WATCHER_TEST_LOG_PATH");
    if (!automatedPath.empty()) {
      const DWORD attributes = GetFileAttributesW(automatedPath.c_str());
      if (attributes != INVALID_FILE_ATTRIBUTES &&
          (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {
        ApplySelectedLogPath(automatedPath);
        DebugLog(L"ChooseLogPath selected path=" + g_state.logPath +
                 L" [automation]");
        return;
      }
      DebugLog(L"ChooseLogPath automation path invalid: " + automatedPath);
    }
  }

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
    DebugLog(L"ChooseLogPath canceled.");
    return;
  }

  ApplySelectedLogPath(filePathBuffer);
  DebugLog(L"ChooseLogPath selected path=" + g_state.logPath);
}

void ShowTrayContextMenu(HWND hwnd) {
  HMENU menu = CreatePopupMenu();
  if (!menu) {
    return;
  }
  const bool showMinutes = g_state.monitorIntervalUseMinutes;
  const double displayIntervalValue = showMinutes
      ? (static_cast<double>(g_state.monitorIntervalMs) / 60000.0)
      : (static_cast<double>(g_state.monitorIntervalMs) / 1000.0);
  const wchar_t* unitLabel = showMinutes ? L"min" : L"s";
  wchar_t intervalMenuText[96] = {};
  StringCchPrintfW(
      intervalMenuText,
      ARRAYSIZE(intervalMenuText),
      L"Set check interval... (current: %.2f %ls)",
      displayIntervalValue,
      unitLabel);
  wchar_t doubleClickMenuText[160] = {};
  StringCchPrintfW(
      doubleClickMenuText,
      ARRAYSIZE(doubleClickMenuText),
      L"Set double-click action... (current: %ls)",
      DoubleClickActionLabel(g_state.doubleClickAction));

  AppendMenuW(menu, MF_STRING, kMenuSetLogPath, L"Set log file path...");
  AppendMenuW(menu, MF_STRING, kMenuSetMonitorInterval, intervalMenuText);
  AppendMenuW(menu, MF_STRING, kMenuSetDoubleClickAction, doubleClickMenuText);
  AppendMenuW(menu, MF_STRING, kMenuOpenAlertMessages, L"Open alert messages...");
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
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

bool HandleCommand(UINT commandId) {
  DebugLog(L"HandleCommand called with id=" + std::to_wstring(commandId));
  switch (commandId) {
    case kMenuSetLogPath:
      ChooseLogPath();
      return true;
    case kMenuSetMonitorInterval:
      PromptAndSetMonitorInterval();
      return true;
    case kMenuSetDoubleClickAction:
      PromptAndSetDoubleClickAction();
      return true;
    case kMenuOpenAlertMessages:
      ShowAlertManagerWindow();
      return true;
    case kMenuOpenLogFolder:
      OpenLogFolder();
      return true;
    case kMenuOpenLogFile:
      OpenLogFile();
      return true;
    case kMenuAcknowledgeAlert:
      AcknowledgeAlert();
      return true;
    case kMenuExit:
      DestroyWindow(g_state.hwnd);
      return true;
    default:
      return false;
  }
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
  g_state.warningIcon = LoadIconW(nullptr, IDI_WARNING);
  g_state.errorIcon = LoadIconW(nullptr, IDI_ERROR);
  if (!g_state.normalIcon || !g_state.warningIcon || !g_state.errorIcon) {
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
  DebugLogWindowMessage(L"MainWindow", message, wParam, lParam);
  if (message == g_state.taskbarCreatedMessage && g_state.taskbarCreatedMessage != 0) {
    AddTrayIcon();
    return 0;
  }

  switch (message) {
    case WM_TIMER:
      if (wParam == kMonitorTimerId) {
        MonitorLogFileOnce();
      } else if (wParam == kBlinkTimerId) {
        if (ShouldBlinkForSeverity(g_state.alertSeverity)) {
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
        HandleCommand(static_cast<UINT>(g_state.doubleClickAction));
      }
      return 0;
    }

    case WM_COMMAND: {
      if (HandleCommand(LOWORD(wParam))) {
        return 0;
      }
      return 0;
    }

    case WM_DESTROY:
      KillTimer(hwnd, kMonitorTimerId);
      KillTimer(hwnd, kBlinkTimerId);
      if (g_state.acknowledgePopupHwnd && IsWindow(g_state.acknowledgePopupHwnd)) {
        DestroyWindow(g_state.acknowledgePopupHwnd);
        g_state.acknowledgePopupHwnd = nullptr;
      }
      if (g_state.alertManagerHwnd && IsWindow(g_state.alertManagerHwnd)) {
        DestroyWindow(g_state.alertManagerHwnd);
        g_state.alertManagerHwnd = nullptr;
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

#ifndef BACKREST_WATCHER_TEST_BUILD
int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
  InitializeDebugModeFromCommandLine();
  bool alreadyRunning = false;
  if (!AcquireSingleInstanceLock(&alreadyRunning)) {
    MessageBoxW(nullptr, L"Failed to initialize single-instance lock.", L"Backrest Watcher", MB_ICONERROR | MB_OK);
    return 1;
  }
  if (alreadyRunning) {
    ReleaseSingleInstanceLock();
    return 0;
  }

  INITCOMMONCONTROLSEX commonControls = {};
  commonControls.dwSize = sizeof(commonControls);
  commonControls.dwICC = ICC_TAB_CLASSES;
  InitCommonControlsEx(&commonControls);

  g_state.taskbarCreatedMessage = RegisterWindowMessageW(L"TaskbarCreated");
  g_state.configPath = ConfigFilePath();
  g_state.ignorePath = IgnoreFilePath();
  g_state.debugLogPath = DebugLogPath();
  LoadLogPathFromConfig();
  ReloadIgnoreListIfChanged(true);

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

  const POINT hiddenOwnerWindowPosition = CenteredWindowPosition(kHiddenOwnerWindowWidth, kHiddenOwnerWindowHeight);
  HWND hwnd = CreateWindowExW(
      0,
      kWindowClassName,
      L"Backrest Watcher",
      WS_OVERLAPPEDWINDOW,
      hiddenOwnerWindowPosition.x,
      hiddenOwnerWindowPosition.y,
      kHiddenOwnerWindowWidth,
      kHiddenOwnerWindowHeight,
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

  DebugLog(L"Process shutting down with exit code " + std::to_wstring(static_cast<int>(message.wParam)));
  ReleaseSingleInstanceLock();
  return static_cast<int>(message.wParam);
}
#endif
