// =====================================================================
// OrbitFM.cpp — v8.4.1
// OrbitFM — A fast, modern file manager for Windows. Win32 only. No external libs.
//
// KEYBOARD SHORTCUTS REFERENCE
// ---------------------------------------------------------------------
// NAVIGATION
//   Up / Down / Left / Right    Move selection in active pane
//   Home / End                  First / last item
//   PageUp / PageDown           Page scroll
//   Enter                       Open file or enter folder
//   Backspace                   Go to parent folder
//   Tab                         Switch active pane (left <-> right)
//   F1                          Move focus to sidebar
//   Ctrl+L                      Edit current path manually
//   Alt+Home                    Go to Home (Quick Access / Recent / This PC)
//
// FILE OPERATIONS
//   F2                          Rename selected
//   Delete                      Send to Recycle Bin
//   Shift+Delete                Delete permanently
//   Ctrl+C                      Copy
//   Ctrl+X                      Cut
//   Ctrl+V                      Paste into current folder
//   Ctrl+A                      Select all
//   F5                          Refresh current folder
//   Ctrl+N                      New folder
//
// TABS
//   Ctrl+T                      New tab in active pane
//   Ctrl+W                      Close current tab
//   Ctrl+Tab                    Next tab
//   Ctrl+Shift+Tab              Previous tab
//
// VIEW
//   Ctrl+H                      Toggle hidden files
//   Ctrl+1                      Sort by Name
//   Ctrl+2                      Sort by Size
//   Ctrl+3                      Sort by Date Modified
//   Ctrl+4                      Sort by Type
// =====================================================================

#define _WIN32_WINNT 0x0601
#define WINVER 0x0601
#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define NOMINMAX

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <initguid.h>
#include <shlobj.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <propkey.h>
#include <propvarutil.h>
#include <propsys.h>
#include <uxtheme.h>
#include <ole2.h>
#include <strsafe.h>

DEFINE_GUID(KFID_Desktop,         0xB4BFCC3A, 0xDB2C, 0x424C, 0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41);
DEFINE_GUID(KFID_Downloads,       0x374DE290, 0x123F, 0x4565, 0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B);
DEFINE_GUID(KFID_Documents,       0xFDD39AD0, 0x238F, 0x46AF, 0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7);
DEFINE_GUID(KFID_Pictures,        0x33E28130, 0x4E1E, 0x4676, 0x83, 0x5A, 0x98, 0x39, 0x5C, 0x3B, 0xC3, 0xBB);
DEFINE_GUID(KFID_Music,           0x4BD8D571, 0x6D19, 0x48D3, 0xBE, 0x97, 0x42, 0x22, 0x20, 0x08, 0x0E, 0x43);
DEFINE_GUID(KFID_Videos,          0x18989B1D, 0x99B5, 0x455B, 0x84, 0x1C, 0xAB, 0x7C, 0x74, 0xE4, 0xDD, 0xFC);
DEFINE_GUID(KFID_RoamingAppData,  0x3EB685DB, 0x65F9, 0x4CF6, 0xA0, 0x3A, 0xE3, 0xEF, 0x65, 0x72, 0x9F, 0x3D);
DEFINE_GUID(KFID_Profile,         0x5E6C858F, 0x0E22, 0x4760, 0x9A, 0xFE, 0xEA, 0x33, 0x17, 0xB6, 0x71, 0x73);

static const PROPERTYKEY PK_MediaDuration =
    { {0x64440490, 0x4C8B, 0x11D1, {0x8B,0x70,0x08,0x00,0x36,0xB1,0x1A,0x03}}, 3 };
static const PROPERTYKEY PK_AudioBitrate =
    { {0x64440490, 0x4C8B, 0x11D1, {0x8B,0x70,0x08,0x00,0x36,0xB1,0x1A,0x03}}, 4 };
static const PROPERTYKEY PK_AudioSampleRate =
    { {0x64440490, 0x4C8B, 0x11D1, {0x8B,0x70,0x08,0x00,0x36,0xB1,0x1A,0x03}}, 5 };
static const PROPERTYKEY PK_MusicArtist =
    { {0x56A3372E, 0xCE9C, 0x11D2, {0x9F,0x0E,0x00,0x60,0x97,0xC6,0x86,0xF6}}, 2 };
static const PROPERTYKEY PK_MusicAlbum =
    { {0x56A3372E, 0xCE9C, 0x11D2, {0x9F,0x0E,0x00,0x60,0x97,0xC6,0x86,0xF6}}, 4 };
static const PROPERTYKEY PK_Title =
    { {0xF29F85E0, 0x4FF9, 0x1068, {0xAB,0x91,0x08,0x00,0x2B,0x27,0xB3,0xD9}}, 2 };

#include <thread>
#include <mutex>
#include <condition_variable>
#include <queue>
#include <unordered_map>
#include <unordered_set>
#include <string>
#include <vector>
#include <algorithm>
#include <fstream>
#include <atomic>
#include <cstdint>
#include <cmath>

#ifdef _MSC_VER
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "propsys.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif

// =====================================================================
// COLORS / SIZES / IDS
// =====================================================================

static COLORREF COL_BG       = RGB(0x57, 0x50, 0x50);
static COLORREF COL_BG_ALT   = RGB(0x4A, 0x44, 0x44);
static COLORREF COL_BORDER   = RGB(0x35, 0x30, 0x30);
static COLORREF COL_SEL      = RGB(0x2A, 0x5A, 0x8A);
static COLORREF COL_HOVER    = RGB(0x62, 0x5C, 0x5C);
static COLORREF COL_TEXT     = RGB(0xEA, 0xE5, 0xE5);
static COLORREF COL_DIM      = RGB(0xA0, 0x98, 0x98);
static COLORREF COL_ACTIVE   = RGB(0x65, 0x5E, 0x5E);
static COLORREF COL_HEADER   = RGB(0x3A, 0x35, 0x35);
static COLORREF COL_ACCENT   = RGB(0x4A, 0x9E, 0xD4);
static COLORREF COL_SIDEBAR  = RGB(0x3E, 0x38, 0x38);

static int SIDEBAR_W = 180;
static int PREVIEW_W = 320;
static const int TAB_H = 30;
static const int CRUMB_H = 26;
static const int TOOLBAR_H = 32;
static const int STATUS_H = 22;
static const int SPLIT_W = 4;
static const int SIDEBAR_ROW_H = 26;
static const int SIDEBAR_HEADER_H = 24;
static const int TAB_X_W = 18;
static const int CRUMB_SEP_W = 16;
static const int PREVIEW_THUMB_PAD = 12;
static const int TOOLBAR_BTN_W = 32;
static const int SEARCH_W = 200;

#define WM_THUMB_READY    (WM_APP + 1)
#define WM_NAVIGATE_PIN   (WM_APP + 2)
#define WM_HASH_READY     (WM_APP + 3)
#define WM_FOLDER_SIZE    (WM_APP + 4)
#define WM_RECENT_CHANGED (WM_APP + 5)
#define WM_AUTOREFRESH    (WM_APP + 6)
#define WM_GIT_READY      (WM_APP + 7)
#define WM_DUPE_READY     (WM_APP + 8)
#define WM_FOLDER_READY   (WM_APP + 9)
#define WM_SEARCH_RESULT  (WM_APP + 10)

#define TIMER_FADE        50   // +paneIndex (50, 51)
#define TIMER_SCROLL      52   // +paneIndex (52, 53)
#define TIMER_ANIM        54   // +paneIndex (54, 55) — hover glow
#define TIMER_SPIN        56   // +paneIndex (56, 57) — loading spinner
#define TIMER_PULSE       58   // +paneIndex (58, 59) — active border pulse
#define TIMER_SIDEBAR     60   // sidebar hover animation

#define ID_LIST            1001
#define ID_PATH_EDIT       1002
#define ID_RENAME_EDIT     1003
#define ID_SEARCH_EDIT     1004
#define ID_BULK_FIND       1005
#define ID_BULK_REPL       1006
#define ID_BULK_PREFIX     1007
#define ID_BULK_LIST       1008
#define ID_BULK_OK         1009
#define ID_BULK_CANCEL     1010
#define ID_BULK_NUMBER     1011

enum {
    IDM_OPEN = 100,
    IDM_OPEN_NEW_TAB,
    IDM_OPEN_OTHER_PANE,
    IDM_PIN_SIDEBAR,
    IDM_UNPIN_SIDEBAR,
    IDM_COPY,
    IDM_CUT,
    IDM_PASTE,
    IDM_DELETE,
    IDM_DELETE_PERM,
    IDM_RENAME,
    IDM_BULK_RENAME,
    IDM_PROPERTIES,
    IDM_NEW_FOLDER,
    IDM_REFRESH,
    IDM_SORT_NAME,
    IDM_SORT_SIZE,
    IDM_SORT_DATE,
    IDM_SORT_TYPE,
    IDM_TOGGLE_HIDDEN,
    IDM_FOLDER_SIZE,
    IDM_COMPUTE_HASH,
    IDM_SELECT_PATTERN,
    IDM_INVERT_SEL,
    IDM_COPY_PATH,
    IDM_COPY_PATH_FULL,
    IDM_COPY_PATH_NAME,
    IDM_COPY_PATH_DIR,
    IDM_OPEN_TERMINAL,
    IDM_ZIP_SELECTED,
    IDM_UNZIP_HERE,
    IDM_UNZIP_TO_FOLDER,
    IDM_UNRAR_HERE,
    IDM_UNRAR_TO_FOLDER,
    IDM_GIT_STATUS,
    IDM_FIND_DUPES,
    IDM_PROC_LOCK,
    IDM_FILE_DIFF,
    IDM_FOLDER_DIFF,
    IDM_SPLIT_FILE,
    IDM_JOIN_FILES,
    IDM_UNDO,
    IDM_SETTINGS,
    IDM_CLIP_HISTORY,
    IDM_QUICK_OPEN,
    IDM_CMD_PALETTE,
    IDM_GRID_VIEW,
    IDM_HEX_VIEW,
    IDM_CONNECT_SERVER,
    IDM_OPEN_WITH,
    IDM_SHELL_MENU
};

// Dialog/control IDs for new features
#define ID_SET_FONT_SIZE  2000
#define ID_SET_SIDEBAR_W  2001
#define ID_SET_PREVIEW_W  2002
#define ID_SET_START_PATH 2003
#define ID_SET_SHOW_HIDDEN 2004
#define ID_SET_AUTO_REFRESH 2005
#define ID_SET_OK         2006
#define ID_SET_CANCEL     2007
#define ID_PAL_LIST       2008
#define ID_PAL_OK         2009
#define ID_PAL_CANCEL     2010
#define ID_PAL_EDIT       2028
#define ID_DUPE_LIST      2011
#define ID_DUPE_DEL       2012
#define ID_DUPE_CANCEL    2013
#define ID_DIFF_LIST      2014
#define ID_DIFF_CANCEL    2015
#define ID_SETT_SECT_BASE 3000   // 3000-3009: section header labels (accent color)
#define ID_SETT_DIM_BASE  3010   // 3010-3019: dim hint labels
#define ID_LOCK_LIST      2016
#define ID_LOCK_KILL      2017
#define ID_LOCK_CANCEL    2018
#define ID_SPLIT_SIZE     2019
#define ID_SPLIT_OK       2020
#define ID_SPLIT_CANCEL   2021
#define ID_QO_EDIT        2022
#define ID_QO_LIST        2023
#define ID_QO_CANCEL      2024
#define ID_CP_LIST        2025
#define ID_CP_OK          2026
#define ID_CP_CANCEL      2027

static const wchar_t* CLASS_MAIN     = L"FMMainWnd";
static const wchar_t* CLASS_SIDEBAR  = L"FMSidebarWnd";
static const wchar_t* CLASS_PANE     = L"FMPaneWnd";
static const wchar_t* CLASS_PREVIEW  = L"FMPreviewWnd";
static const wchar_t* CLASS_STATUS   = L"FMStatusWnd";
static const wchar_t* CLASS_PALETTE  = L"FMPaletteWnd";
static const wchar_t* CLASS_CMDPAL   = L"FMCmdPalWnd";
static const wchar_t* CLASS_QUICKOPEN= L"FMQuickOpenWnd";

enum Theme { THEME_LIGHT = 0, THEME_DARK = 1 };

struct ToolbarBtn { RECT rect; const wchar_t* glyph; int id; bool enabled; };

// =====================================================================
// DATA STRUCTURES
// =====================================================================

struct FileEntry {
    std::wstring name;
    std::wstring type;
    uint64_t size;
    FILETIME modified;
    DWORD attr;
    bool isDir() const { return (attr & FILE_ATTRIBUTE_DIRECTORY) != 0; }
};

struct PaneTab {
    std::wstring path;
    std::vector<FileEntry> entries;
    std::vector<FileEntry> filteredEntries;
    int sortCol = 0;
    bool sortAsc = true;
    int focused = 0;
    std::vector<std::wstring> history;
    int histPos = -1;
    std::wstring search;
};

struct HomeItem { RECT rect; std::wstring path; };

struct Pane {
    HWND hwnd = nullptr;
    HWND hwndList = nullptr;
    HWND hwndPathEdit = nullptr;
    HWND hwndSearch = nullptr;
    std::vector<PaneTab> tabs;
    int activeTab = 0;
    int paneIndex = 0;
    int hoverTab = -1;
    int hoverCrumb = -1;
    int hoverBtn = -1;
    std::vector<RECT> tabRects;
    std::vector<RECT> tabCloseRects;
    std::vector<RECT> crumbRects;
    std::vector<std::wstring> crumbSegs;
    std::vector<std::wstring> crumbPaths;
    std::wstring crumbCached;
    std::vector<ToolbarBtn> tbButtons;
    float scrollVelocity = 0.f;
    int   fadeAlpha = 255;
    float tbHoverProg[8] = {};   // per-button hover glow 0.0→1.0
    bool  isLoading = false;
    int   spinFrame = 0;
    // Home view
    std::vector<HomeItem> homeItems;
    int homeHover = -1;
};

struct AppCtx {
    HINSTANCE hInst = nullptr;
    HWND hwndMain = nullptr;
    HWND hwndSidebar = nullptr;
    HWND hwndPreview = nullptr;
    Pane panes[2];
    int activePane = 0;
    bool showHidden = false;

    std::vector<std::wstring> pinned;
    int sidebarHover = -1;
    std::vector<RECT> sidebarRects;

    std::vector<std::wstring> clipboardFiles;
    std::unordered_set<std::wstring> clipboardCutSet;
    bool clipboardCut = false;

    HFONT hFont = nullptr;
    HFONT hFontBold = nullptr;
    HBRUSH brBg = nullptr;
    HBRUSH brBgAlt = nullptr;
    HBRUSH brSel = nullptr;
    HBRUSH brHeader = nullptr;

    std::thread thumbThread;
    std::mutex thumbMu;
    std::condition_variable thumbCv;
    std::queue<std::wstring> thumbQueue;
    std::unordered_map<std::wstring, HBITMAP> thumbCache;
    std::unordered_map<std::wstring, std::wstring> metaCache;
    std::atomic<bool> thumbShutdown{false};

    std::wstring previewPath;

    bool dragging = false;
    bool dragArmed = false;
    std::vector<std::wstring> dragFiles;
    int dragSourcePane = -1;
    POINT dragStart{0, 0};

    HACCEL hAccel = nullptr;

    HIMAGELIST sysImageList = nullptr;
    int folderIconIdx = -1;
    std::unordered_map<std::wstring, int> iconByExt;
    std::mutex iconMu;

    IContextMenu2* shellCM2 = nullptr;
    IContextMenu3* shellCM3 = nullptr;

    int theme = THEME_DARK;
    bool showPane2 = false;
    HWND hwndStatus = nullptr;
    HWND hwndSettings = nullptr;
    std::wstring statusText;

    int sidebarW = 180;
    int previewW = 320;
    int splitDrag = -1;
    POINT splitDragStart{0, 0};
    int splitDragOrig = 0;

    std::vector<std::wstring> recent;

    std::thread hashThread;
    std::mutex hashMu;
    std::condition_variable hashCv;
    std::queue<std::wstring> hashQueue;
    std::unordered_map<std::wstring, std::wstring> hashCache;
    std::atomic<bool> hashShutdown{false};

    std::thread sizeThread;
    std::atomic<bool> sizeRunning{false};
    std::atomic<bool> sizeCancel{false};
    std::atomic<uint64_t> sizeBytes{0};
    std::atomic<uint64_t> sizeFiles{0};
    std::wstring sizeRoot;

    // --- New features ---

    // Settings
    int settingFontSize = 10;
    std::wstring settingStartPath;
    bool settingAutoRefresh = true;
    std::unordered_map<std::wstring, std::wstring> customAssoc; // ext->app

    // View mode per pane (0=details, 1=grid)
    int viewMode[2] = {0, 0};

    // Color tags: path -> COLORREF
    std::unordered_map<std::wstring, COLORREF> colorTags;

    // Multi-clipboard history (last 8 entries)
    struct ClipEntry { std::vector<std::wstring> files; bool cut; };
    std::vector<ClipEntry> clipHistory;

    // Bookmarks Ctrl+Shift+1..9 set, Ctrl+5..9 jump
    std::wstring bookmarks[9];

    // Undo stack: type 0=move, 1=copy, 2=rename, 3=mkdir
    struct UndoOp { int type; std::wstring src; std::wstring dst; };
    std::vector<UndoOp> undoStack;

    // Git status cache: filename -> status char
    std::unordered_map<std::wstring, wchar_t> gitStatus;
    std::thread gitThread;
    std::atomic<bool> gitRunning{false};

    // Drives sidebar + OneDrive
    std::vector<std::wstring> drives;
    std::wstring oneDrivePath;
    int sidebarDriveBase = 0;
    int sidebarHomeIdx   = 0;
    int sidebarPinnedBase = 1;

    // Monospace font for text preview
    HFONT hFontMono = nullptr;

    // Async folder load: serial per pane prevents stale results
    int paneLoadSerial[2] = {0, 0};

    // Background recursive search
    std::atomic<bool> searchBgStop{false};
    std::thread searchBgThread;

    // Auto-refresh watch threads (one per pane)
    std::thread watchThreads[2];
    std::atomic<bool> watchStop{false};

    // Hex view toggle per pane
    bool hexView[2] = {false, false};

    // Duplicate finder results
    std::vector<std::vector<std::wstring>> dupeGroups;
    std::thread dupeThread;
    std::atomic<bool> dupeRunning{false};

    // Quick-open results
    std::vector<std::wstring> quickOpenResults;
    std::thread quickOpenThread;
    std::atomic<bool> quickOpenStop{false};
    HWND hwndQuickOpen = nullptr;
    HWND hwndCmdPal = nullptr;

    // Sidebar hover animation
    float sidebarHoverProg = 0.f;
    float sidebarPrevProg  = 0.f;
    int   sidebarPrevHover = -1;

    // Tab tear-off drag state
    bool tabDragging = false;
    int  tabDragPane = -1;
    int  tabDragIdx  = -1;
    bool tabTearOut  = false;   // true once cursor leaves window
    POINT tabDragStart{0,0};
};

static AppCtx g;

// =====================================================================
// UTILITY HELPERS
// =====================================================================

static std::wstring KnownFolderById(const GUID& id) {
    PWSTR p = nullptr;
    std::wstring r;
    if (SUCCEEDED(SHGetKnownFolderPath(id, 0, nullptr, &p)) && p) r = p;
    if (p) CoTaskMemFree(p);
    return r;
}

static std::wstring ConfigDir() {
    std::wstring r = KnownFolderById(KFID_RoamingAppData);
    if (r.empty()) {
        wchar_t buf[MAX_PATH];
        if (SUCCEEDED(SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, 0, buf))) r = buf;
    }
    if (r.empty()) return L"";
    r += L"\\OrbitFM";
    CreateDirectoryW(r.c_str(), nullptr);
    return r;
}

static std::wstring PinsFile() {
    auto d = ConfigDir();
    return d.empty() ? L"" : d + L"\\pins.txt";
}

static std::wstring RecentFile() {
    auto d = ConfigDir();
    return d.empty() ? L"" : d + L"\\recent.txt";
}

static std::wstring ThemeFile() {
    auto d = ConfigDir();
    return d.empty() ? L"" : d + L"\\theme.txt";
}

static std::wstring SettingsFile() {
    auto d = ConfigDir();
    return d.empty() ? L"" : d + L"\\settings.ini";
}

static std::wstring SessionFile() {
    auto d = ConfigDir();
    return d.empty() ? L"" : d + L"\\session.txt";
}

static void ApplyThemePalette(int /*theme*/) {
    COL_BG      = RGB(0x57, 0x50, 0x50);
    COL_BG_ALT  = RGB(0x4A, 0x44, 0x44);
    COL_BORDER  = RGB(0x2D, 0x28, 0x28);
    COL_SEL     = RGB(0x2A, 0x5A, 0x8A);
    COL_HOVER   = RGB(0x62, 0x5C, 0x5C);
    COL_TEXT    = RGB(0xEA, 0xE5, 0xE5);
    COL_DIM     = RGB(0xA0, 0x98, 0x98);
    COL_ACTIVE  = RGB(0x65, 0x5E, 0x5E);
    COL_HEADER  = RGB(0x3A, 0x35, 0x35);
    COL_ACCENT  = RGB(0x4A, 0x9E, 0xD4);
    COL_SIDEBAR = RGB(0x3E, 0x38, 0x38);
}

static void LoadTheme() {
    g.theme = THEME_DARK;
    ApplyThemePalette(THEME_DARK);
}

static void SaveTheme() {}

static std::wstring ParentPath(const std::wstring& p) {
    if (p.size() < 3) return p;
    size_t idx = p.find_last_of(L'\\');
    if (idx == std::wstring::npos) return p;
    if (idx == 2 && p.size() > 3) return p.substr(0, 3);
    if (idx < 2) return p;
    return p.substr(0, idx);
}

static std::wstring NameFromPath(const std::wstring& p) {
    if (p.size() == 3 && p[1] == L':') return p;
    size_t idx = p.find_last_of(L'\\');
    if (idx == std::wstring::npos) return p;
    return p.substr(idx + 1);
}

static std::wstring CombinePath(const std::wstring& a, const std::wstring& b) {
    if (a.empty()) return b;
    if (b.empty()) return a;
    if (a.back() == L'\\') return a + b;
    return a + L"\\" + b;
}

static std::wstring NormalizePath(std::wstring p) {
    while (!p.empty() && (p.back() == L' ' || p.back() == L'\t')) p.pop_back();
    while (!p.empty() && p.front() == L' ') p.erase(p.begin());
    if (p.size() > 3 && p.back() == L'\\') p.pop_back();
    return p;
}

static std::wstring ExtFromName(const std::wstring& n) {
    size_t i = n.find_last_of(L'.');
    if (i == std::wstring::npos || i == 0) return L"";
    std::wstring e = n.substr(i + 1);
    for (auto& c : e) c = (wchar_t)towupper(c);
    return e;
}

static bool IsImageExt(const std::wstring& e) {
    return e == L"JPG" || e == L"JPEG" || e == L"PNG" || e == L"BMP" ||
           e == L"GIF" || e == L"WEBP" || e == L"TIF" || e == L"TIFF" || e == L"ICO";
}
static bool IsVideoExt(const std::wstring& e) {
    return e == L"MP4" || e == L"MKV" || e == L"AVI" || e == L"MOV" ||
           e == L"WMV" || e == L"WEBM" || e == L"FLV" || e == L"M4V" || e == L"3GP";
}
static bool IsAudioExt(const std::wstring& e) {
    return e == L"MP3" || e == L"WAV" || e == L"FLAC" || e == L"OGG" ||
           e == L"AAC" || e == L"M4A" || e == L"WMA" || e == L"OPUS";
}
static bool IsTextExt(const std::wstring& e) {
    static const wchar_t* exts[] = {
        L"TXT",L"MD",L"LOG",L"JSON",L"XML",L"CSV",L"PY",L"JS",L"TS",
        L"CPP",L"H",L"C",L"CS",L"INI",L"YAML",L"YML",L"TOML",
        L"BAT",L"PS1",L"SH",L"CMAKE",L"RC",L"CFG",L"CONF",L"GITIGNORE",nullptr
    };
    for (int i = 0; exts[i]; i++) if (e == exts[i]) return true;
    return false;
}

// Result packet posted from async folder-load thread to pane window
struct FolderResult {
    int paneIdx;
    int serial;
    std::wstring path;
    std::vector<FileEntry> entries;
};

static bool DirExists(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_DIRECTORY);
}

static void FormatSize(uint64_t bytes, wchar_t* buf, size_t cch) {
    const wchar_t* units[] = { L"B", L"KB", L"MB", L"GB", L"TB" };
    double v = (double)bytes;
    int u = 0;
    while (v >= 1024.0 && u < 4) { v /= 1024.0; u++; }
    if (u == 0) StringCchPrintfW(buf, cch, L"%llu B", (unsigned long long)bytes);
    else StringCchPrintfW(buf, cch, L"%.1f %s", v, units[u]);
}

static void FormatTimeFt(const FILETIME& ft, wchar_t* buf, size_t cch) {
    SYSTEMTIME st;
    FILETIME local;
    if (!FileTimeToLocalFileTime(&ft, &local) || !FileTimeToSystemTime(&local, &st)) {
        StringCchCopyW(buf, cch, L"");
        return;
    }
    StringCchPrintfW(buf, cch, L"%04u-%02u-%02u  %02u:%02u",
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute);
}

static std::vector<std::wstring> SplitPath(const std::wstring& p) {
    std::vector<std::wstring> out;
    if (p.empty()) return out;
    if (p.size() >= 2 && p[1] == L':') {
        out.push_back(p.substr(0, 2));
        size_t i = (p.size() > 2 && p[2] == L'\\') ? 3 : 2;
        while (i < p.size()) {
            size_t j = p.find(L'\\', i);
            if (j == std::wstring::npos) { out.push_back(p.substr(i)); break; }
            if (j > i) out.push_back(p.substr(i, j - i));
            i = j + 1;
        }
    }
    return out;
}

static std::wstring TabLabel(const std::wstring& path) {
    if (path == L"::home::") return L"Home";
    if (path.size() == 2 && path[1] == L':') return path;
    if (path.size() == 3 && path[1] == L':') return path;
    return NameFromPath(path);
}

static int MeasureText(HDC hdc, const std::wstring& s) {
    SIZE sz{};
    GetTextExtentPoint32W(hdc, s.c_str(), (int)s.size(), &sz);
    return sz.cx;
}

// =====================================================================
// PIN LOAD/SAVE
// =====================================================================

static void LoadPins() {
    g.pinned.clear();
    auto file = PinsFile();
    if (!file.empty()) {
        HANDLE h = CreateFileW(file.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h != INVALID_HANDLE_VALUE) {
            LARGE_INTEGER sz{};
            GetFileSizeEx(h, &sz);
            if (sz.QuadPart > 0 && sz.QuadPart < 1024 * 1024) {
                std::vector<char> buf((size_t)sz.QuadPart);
                DWORD got = 0;
                ReadFile(h, buf.data(), (DWORD)buf.size(), &got, nullptr);
                std::wstring all;
                size_t off = 0;
                if (got >= 2 && (unsigned char)buf[0] == 0xFF && (unsigned char)buf[1] == 0xFE) {
                    off = 2;
                    all.assign((wchar_t*)(buf.data() + off), (got - off) / 2);
                } else {
                    int wn = MultiByteToWideChar(CP_UTF8, 0, buf.data(), got, nullptr, 0);
                    if (wn > 0) {
                        all.resize((size_t)wn);
                        MultiByteToWideChar(CP_UTF8, 0, buf.data(), got, all.data(), wn);
                    }
                }
                size_t s = 0;
                for (size_t i = 0; i <= all.size(); i++) {
                    if (i == all.size() || all[i] == L'\n' || all[i] == L'\r') {
                        if (i > s) {
                            std::wstring line = all.substr(s, i - s);
                            if (!line.empty() && line.back() == L'\r') line.pop_back();
                            if (!line.empty() && DirExists(line)) g.pinned.push_back(line);
                        }
                        s = i + 1;
                    }
                }
            }
            CloseHandle(h);
        }
    }
    if (g.pinned.empty()) {
        const GUID defs[] = {
            KFID_Desktop, KFID_Downloads, KFID_Documents,
            KFID_Pictures, KFID_Music, KFID_Videos
        };
        for (const auto& id : defs) {
            auto p = KnownFolderById(id);
            if (!p.empty() && DirExists(p)) g.pinned.push_back(p);
        }
    }
}

static void SavePins() {
    auto file = PinsFile();
    if (file.empty()) return;
    HANDLE h = CreateFileW(file.c_str(), GENERIC_WRITE, 0, nullptr,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    const unsigned char bom[2] = { 0xFF, 0xFE };
    DWORD wrote = 0;
    WriteFile(h, bom, 2, &wrote, nullptr);
    for (auto& p : g.pinned) {
        WriteFile(h, p.data(), (DWORD)(p.size() * sizeof(wchar_t)), &wrote, nullptr);
        wchar_t nl = L'\n';
        WriteFile(h, &nl, sizeof(wchar_t), &wrote, nullptr);
    }
    CloseHandle(h);
}

// =====================================================================
// FOLDER ENUMERATION + SORT
// =====================================================================

static std::vector<FileEntry> EnumerateFolder(const std::wstring& path, bool showHidden) {
    std::vector<FileEntry> result;
    if (path.empty()) return result;
    std::wstring pattern = path;
    if (pattern.back() != L'\\') pattern += L"\\";
    pattern += L"*";
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return result;
    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        if (!showHidden && (fd.dwFileAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM))) continue;
        FileEntry e;
        e.name = fd.cFileName;
        e.size = ((uint64_t)fd.nFileSizeHigh << 32) | fd.nFileSizeLow;
        e.modified = fd.ftLastWriteTime;
        e.attr = fd.dwFileAttributes;
        if (e.isDir()) {
            e.type = L"Folder";
            e.size = 0;
        } else {
            std::wstring x = ExtFromName(e.name);
            e.type = x.empty() ? L"File" : x + L" File";
        }
        result.push_back(std::move(e));
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return result;
}

static int CmpStrI(const std::wstring& a, const std::wstring& b) {
    return _wcsicmp(a.c_str(), b.c_str());
}

static void SortEntries(std::vector<FileEntry>& v, int col, bool asc) {
    std::sort(v.begin(), v.end(), [&](const FileEntry& a, const FileEntry& b) {
        if (a.isDir() != b.isDir()) return a.isDir();
        int c = 0;
        switch (col) {
            case 0: c = CmpStrI(a.name, b.name); break;
            case 1: c = (a.size < b.size) ? -1 : (a.size > b.size ? 1 : CmpStrI(a.name, b.name)); break;
            case 2: c = CompareFileTime(&a.modified, &b.modified); if (c == 0) c = CmpStrI(a.name, b.name); break;
            case 3: c = CmpStrI(a.type, b.type); if (c == 0) c = CmpStrI(a.name, b.name); break;
        }
        return asc ? (c < 0) : (c > 0);
    });
}

// =====================================================================
// THUMBNAIL CACHE / WORKER THREAD
// =====================================================================

static HBITMAP MakeShellThumb(const std::wstring& path, int sz) {
    HBITMAP hbmp = nullptr;
    IShellItemImageFactory* factory = nullptr;
    HRESULT hr = SHCreateItemFromParsingName(path.c_str(), nullptr, IID_PPV_ARGS(&factory));
    if (SUCCEEDED(hr) && factory) {
        SIZE s{ sz, sz };
        hr = factory->GetImage(s, SIIGBF_THUMBNAILONLY | SIIGBF_BIGGERSIZEOK, &hbmp);
        if (FAILED(hr) || !hbmp) {
            if (hbmp) { DeleteObject(hbmp); hbmp = nullptr; }
            hr = factory->GetImage(s, SIIGBF_RESIZETOFIT, &hbmp);
            if (FAILED(hr) && hbmp) { DeleteObject(hbmp); hbmp = nullptr; }
        }
        factory->Release();
    }
    return hbmp;
}

static std::wstring GetMediaMeta(const std::wstring& path) {
    std::wstring out;
    IShellItem2* psi = nullptr;
    HRESULT hr = SHCreateItemFromParsingName(path.c_str(), nullptr, IID_PPV_ARGS(&psi));
    if (FAILED(hr) || !psi) return out;
    auto getStr = [&](REFPROPERTYKEY k)->std::wstring {
        PWSTR s = nullptr; std::wstring r;
        if (SUCCEEDED(psi->GetString(k, &s)) && s) { r = s; CoTaskMemFree(s); }
        return r;
    };
    auto getU64 = [&](REFPROPERTYKEY k)->uint64_t {
        ULONGLONG v = 0;
        if (SUCCEEDED(psi->GetUInt64(k, &v))) return v;
        return 0;
    };
    auto getU32 = [&](REFPROPERTYKEY k)->uint32_t {
        ULONG v = 0;
        if (SUCCEEDED(psi->GetUInt32(k, &v))) return (uint32_t)v;
        return 0;
    };
    uint64_t dur = getU64(PK_MediaDuration);
    if (dur > 0) {
        wchar_t buf[64];
        uint64_t totalSec = dur / 10000000ULL;
        uint64_t hh = totalSec / 3600, mm = (totalSec / 60) % 60, ss = totalSec % 60;
        if (hh > 0) StringCchPrintfW(buf, 64, L"Duration: %llu:%02llu:%02llu", hh, mm, ss);
        else        StringCchPrintfW(buf, 64, L"Duration: %llu:%02llu", mm, ss);
        out += buf; out += L"\n";
    }
    std::wstring title = getStr(PK_Title);
    std::wstring artist = getStr(PK_MusicArtist);
    std::wstring album = getStr(PK_MusicAlbum);
    uint32_t bitrate = getU32(PK_AudioBitrate);
    uint32_t sr = getU32(PK_AudioSampleRate);
    if (!title.empty())  { out += L"Title: ";  out += title;  out += L"\n"; }
    if (!artist.empty()) { out += L"Artist: "; out += artist; out += L"\n"; }
    if (!album.empty())  { out += L"Album: ";  out += album;  out += L"\n"; }
    if (bitrate > 0) {
        wchar_t b[32]; StringCchPrintfW(b, 32, L"Bitrate: %u kbps\n", bitrate / 1000); out += b;
    }
    if (sr > 0) {
        wchar_t b[32]; StringCchPrintfW(b, 32, L"Sample Rate: %u Hz\n", sr); out += b;
    }
    psi->Release();
    return out;
}

static void ThumbWorker() {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    while (!g.thumbShutdown.load()) {
        std::wstring path;
        {
            std::unique_lock<std::mutex> lk(g.thumbMu);
            g.thumbCv.wait(lk, [] { return g.thumbShutdown.load() || !g.thumbQueue.empty(); });
            if (g.thumbShutdown.load()) break;
            path = g.thumbQueue.front();
            g.thumbQueue.pop();
        }
        std::wstring ext = ExtFromName(path);
        bool isAudio = IsAudioExt(ext);
        bool wantImg = IsImageExt(ext) || IsVideoExt(ext) || isAudio;
        if (!wantImg) continue;

        HBITMAP bmp = MakeShellThumb(path, 256);
        std::wstring meta;
        if (isAudio || IsVideoExt(ext)) meta = GetMediaMeta(path);

        {
            std::lock_guard<std::mutex> lk(g.thumbMu);
            auto it = g.thumbCache.find(path);
            if (it != g.thumbCache.end() && it->second) DeleteObject(it->second);
            g.thumbCache[path] = bmp;
            g.metaCache[path] = meta;
        }
        if (g.hwndPreview) {
            wchar_t* dup = (wchar_t*)CoTaskMemAlloc((path.size() + 1) * sizeof(wchar_t));
            if (dup) {
                StringCchCopyW(dup, path.size() + 1, path.c_str());
                PostMessageW(g.hwndPreview, WM_THUMB_READY, 0, (LPARAM)dup);
            }
        }
    }
    CoUninitialize();
}

static void RequestThumb(const std::wstring& path) {
    {
        std::lock_guard<std::mutex> lk(g.thumbMu);
        if (g.thumbCache.find(path) != g.thumbCache.end()) {
            if (g.hwndPreview) InvalidateRect(g.hwndPreview, nullptr, FALSE);
            return;
        }
        std::queue<std::wstring> empty;
        std::swap(g.thumbQueue, empty);
        g.thumbQueue.push(path);
    }
    g.thumbCv.notify_one();
}

static void StartThumbThread() {
    g.thumbThread = std::thread(ThumbWorker);
}

static void StopThumbThread() {
    g.thumbShutdown.store(true);
    g.thumbCv.notify_all();
    if (g.thumbThread.joinable()) g.thumbThread.join();
    for (auto& kv : g.thumbCache) if (kv.second) DeleteObject(kv.second);
    g.thumbCache.clear();
    g.metaCache.clear();
}

static void RebuildBrushes() {
    if (g.brBg)     DeleteObject(g.brBg);
    if (g.brBgAlt)  DeleteObject(g.brBgAlt);
    if (g.brSel)    DeleteObject(g.brSel);
    if (g.brHeader) DeleteObject(g.brHeader);
    g.brBg     = CreateSolidBrush(COL_BG);
    g.brBgAlt  = CreateSolidBrush(COL_BG_ALT);
    g.brSel    = CreateSolidBrush(COL_SEL);
    g.brHeader = CreateSolidBrush(COL_HEADER);
}

static void ApplyTheme(int theme) {
    g.theme = theme;
    ApplyThemePalette(theme);
    RebuildBrushes();
    for (int i = 0; i < 2; i++) {
        if (g.panes[i].hwndList) {
            ListView_SetBkColor(g.panes[i].hwndList, COL_BG);
            ListView_SetTextBkColor(g.panes[i].hwndList, COL_BG);
            ListView_SetTextColor(g.panes[i].hwndList, COL_TEXT);
            InvalidateRect(g.panes[i].hwndList, nullptr, TRUE);
            HWND hHdr = ListView_GetHeader(g.panes[i].hwndList);
            if (hHdr) InvalidateRect(hHdr, nullptr, TRUE);
        }
        if (g.panes[i].hwnd) InvalidateRect(g.panes[i].hwnd, nullptr, TRUE);
    }
    if (g.hwndSidebar) InvalidateRect(g.hwndSidebar, nullptr, TRUE);
    if (g.hwndPreview) InvalidateRect(g.hwndPreview, nullptr, TRUE);
    if (g.hwndStatus)  InvalidateRect(g.hwndStatus,  nullptr, TRUE);
    if (g.hwndMain)    InvalidateRect(g.hwndMain,    nullptr, TRUE);
}

static void ToggleTheme() {
    ApplyTheme(g.theme == THEME_DARK ? THEME_LIGHT : THEME_DARK);
    SaveTheme();
}

// =====================================================================
// FILE OPERATIONS
// =====================================================================

static std::vector<wchar_t> MakeDoubleNullList(const std::vector<std::wstring>& v) {
    std::vector<wchar_t> b;
    for (auto& s : v) {
        b.insert(b.end(), s.begin(), s.end());
        b.push_back(0);
    }
    b.push_back(0);
    if (b.size() == 1) b.push_back(0);
    return b;
}

static bool ShellOp(HWND hwnd, UINT op, const std::vector<std::wstring>& from,
                    const std::wstring& to, bool allowUndo) {
    auto src = MakeDoubleNullList(from);
    std::wstring dest = to;
    if (!dest.empty()) dest.push_back(0);
    SHFILEOPSTRUCTW s{};
    s.hwnd = hwnd;
    s.wFunc = op;
    s.pFrom = src.data();
    s.pTo = dest.empty() ? nullptr : dest.c_str();
    s.fFlags = (allowUndo ? FOF_ALLOWUNDO : 0) | FOF_NOCONFIRMMKDIR;
    int r = SHFileOperationW(&s);
    return r == 0 && !s.fAnyOperationsAborted;
}

static bool DeleteFiles(HWND hwnd, const std::vector<std::wstring>& files, bool permanent) {
    if (files.empty()) return false;
    return ShellOp(hwnd, FO_DELETE, files, L"", !permanent);
}

static bool CopyFiles(HWND hwnd, const std::vector<std::wstring>& files, const std::wstring& destDir) {
    if (files.empty() || destDir.empty()) return false;
    return ShellOp(hwnd, FO_COPY, files, destDir, true);
}

static bool MoveFiles(HWND hwnd, const std::vector<std::wstring>& files, const std::wstring& destDir) {
    if (files.empty() || destDir.empty()) return false;
    return ShellOp(hwnd, FO_MOVE, files, destDir, true);
}

static bool RenameFile(HWND hwnd, const std::wstring& oldPath, const std::wstring& newName) {
    std::wstring dir = ParentPath(oldPath);
    std::wstring newPath = CombinePath(dir, newName);
    if (newPath == oldPath) return true;
    if (MoveFileExW(oldPath.c_str(), newPath.c_str(), MOVEFILE_COPY_ALLOWED)) return true;
    SHFILEOPSTRUCTW s{};
    std::vector<wchar_t> from; from.insert(from.end(), oldPath.begin(), oldPath.end());
    from.push_back(0); from.push_back(0);
    std::vector<wchar_t> to; to.insert(to.end(), newPath.begin(), newPath.end());
    to.push_back(0); to.push_back(0);
    s.hwnd = hwnd; s.wFunc = FO_RENAME; s.pFrom = from.data(); s.pTo = to.data();
    s.fFlags = FOF_ALLOWUNDO;
    return SHFileOperationW(&s) == 0;
}

static void ShowProperties(HWND hwnd, const std::wstring& path) {
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_INVOKEIDLIST | SEE_MASK_FLAG_NO_UI;
    sei.hwnd = hwnd;
    sei.lpVerb = L"properties";
    sei.lpFile = path.c_str();
    sei.nShow = SW_SHOW;
    ShellExecuteExW(&sei);
}

static void OpenWithDefault(const std::wstring& path) {
    ShellExecuteW(nullptr, L"open", path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

static bool MakeNewFolder(const std::wstring& parent, std::wstring& outName) {
    for (int i = 1; i < 1000; ++i) {
        std::wstring name = (i == 1) ? std::wstring(L"New Folder") :
            (std::wstring(L"New Folder (") + std::to_wstring(i) + L")");
        std::wstring p = CombinePath(parent, name);
        if (CreateDirectoryW(p.c_str(), nullptr)) {
            outName = name;
            return true;
        }
        if (GetLastError() != ERROR_ALREADY_EXISTS) return false;
    }
    return false;
}

// =====================================================================
// RECENT FILES
// =====================================================================

static void LoadRecent() {
    g.recent.clear();
    auto file = RecentFile();
    if (file.empty()) return;
    HANDLE h = CreateFileW(file.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    LARGE_INTEGER sz{};
    GetFileSizeEx(h, &sz);
    if (sz.QuadPart > 0 && sz.QuadPart < 512 * 1024) {
        std::vector<char> buf((size_t)sz.QuadPart);
        DWORD got = 0;
        ReadFile(h, buf.data(), (DWORD)buf.size(), &got, nullptr);
        std::wstring all;
        size_t off = 0;
        if (got >= 2 && (unsigned char)buf[0] == 0xFF && (unsigned char)buf[1] == 0xFE) {
            off = 2; all.assign((wchar_t*)(buf.data() + off), (got - off) / 2);
        } else {
            int wn = MultiByteToWideChar(CP_UTF8, 0, buf.data(), got, nullptr, 0);
            if (wn > 0) { all.resize(wn); MultiByteToWideChar(CP_UTF8, 0, buf.data(), got, all.data(), wn); }
        }
        size_t s = 0;
        for (size_t i = 0; i <= all.size(); i++) {
            if (i == all.size() || all[i] == L'\n' || all[i] == L'\r') {
                if (i > s) {
                    auto ln = all.substr(s, i - s);
                    while (!ln.empty() && (ln.back() == L'\r' || ln.back() == L'\n')) ln.pop_back();
                    if (!ln.empty()) g.recent.push_back(ln);
                }
                s = i + 1;
            }
        }
    }
    CloseHandle(h);
}

static void SaveRecent() {
    auto file = RecentFile();
    if (file.empty()) return;
    HANDLE h = CreateFileW(file.c_str(), GENERIC_WRITE, 0, nullptr,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    const unsigned char bom[2] = { 0xFF, 0xFE };
    DWORD wrote = 0;
    WriteFile(h, bom, 2, &wrote, nullptr);
    for (auto& p : g.recent) {
        WriteFile(h, p.data(), (DWORD)(p.size() * sizeof(wchar_t)), &wrote, nullptr);
        wchar_t nl = L'\n';
        WriteFile(h, &nl, sizeof(wchar_t), &wrote, nullptr);
    }
    CloseHandle(h);
}

static void AddRecent(const std::wstring& path) {
    g.recent.erase(std::remove_if(g.recent.begin(), g.recent.end(),
        [&](const std::wstring& s){ return CmpStrI(s, path) == 0; }), g.recent.end());
    g.recent.insert(g.recent.begin(), path);
    if (g.recent.size() > 30) g.recent.resize(30);
    SaveRecent();
    if (g.hwndSidebar) InvalidateRect(g.hwndSidebar, nullptr, TRUE);
}

// =====================================================================
// SEARCH FILTER
// =====================================================================

static bool MatchesSearch(const FileEntry& e, const std::wstring& q) {
    if (q.empty()) return true;
    // Support simple "size>X" "date>X" later — for now, name substring (case-insensitive)
    std::wstring uq = q, un = e.name;
    for (auto& c : uq) c = (wchar_t)towupper(c);
    for (auto& c : un) c = (wchar_t)towupper(c);
    if (un.find(uq) != std::wstring::npos) return true;
    // also match extension
    std::wstring ux = ExtFromName(e.name);
    return CmpStrI(ux, q) == 0;
}

static void ApplySearchFilter(Pane& p) {
    if (p.tabs.empty()) return;
    auto& t = p.tabs[p.activeTab];
    t.filteredEntries.clear();
    if (t.search.empty()) return;
    for (auto& e : t.entries)
        if (MatchesSearch(e, t.search)) t.filteredEntries.push_back(e);
}

// =====================================================================
// HASH (MD5 + SHA256) — small pure-C implementation
// =====================================================================

static void MD5Block(uint32_t state[4], const uint8_t block[64]);
static void SHA256Block(uint32_t state[8], const uint8_t block[64]);

static std::wstring ComputeHash(const std::wstring& path) {
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
        nullptr, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
    if (h == INVALID_HANDLE_VALUE) return L"";

    // MD5 state
    uint32_t md5s[4] = {0x67452301u,0xEFCDAB89u,0x98BADCFEu,0x10325476u};
    // SHA256 state
    uint32_t sha[8] = {0x6a09e667u,0xbb67ae85u,0x3c6ef372u,0xa54ff53au,
                       0x510e527fu,0x9b05688cu,0x1f83d9abu,0x5be0cd19u};

    uint8_t ibuf[64];
    uint8_t md5buf[128]={}, sha256buf[128]={};
    uint64_t totalBytes = 0;
    uint32_t md5rem = 0, sharem = 0;
    DWORD got;
    while (ReadFile(h, ibuf, 64, &got, nullptr) && got > 0) {
        totalBytes += got;
        // feed MD5
        uint32_t m5off = 0;
        while (m5off < got) {
            uint32_t take = std::min((uint32_t)(got - m5off), 64u - md5rem);
            memcpy(md5buf + md5rem, ibuf + m5off, take);
            md5rem += take; m5off += take;
            if (md5rem == 64) { MD5Block(md5s, md5buf); md5rem = 0; }
        }
        // feed SHA256
        uint32_t s2off = 0;
        while (s2off < got) {
            uint32_t take = std::min((uint32_t)(got - s2off), 64u - sharem);
            memcpy(sha256buf + sharem, ibuf + s2off, take);
            sharem += take; s2off += take;
            if (sharem == 64) { SHA256Block(sha, sha256buf); sharem = 0; }
        }
    }
    CloseHandle(h);

    // MD5 finalize
    {
        uint64_t bitLen = totalBytes * 8;
        md5buf[md5rem++] = 0x80;
        if (md5rem > 56) { memset(md5buf+md5rem,0,64-md5rem); MD5Block(md5s,md5buf); md5rem=0; }
        memset(md5buf+md5rem,0,56-md5rem);
        // little-endian bit length
        for (int i=0;i<8;i++) md5buf[56+i]=(uint8_t)(bitLen>>(8*i));
        MD5Block(md5s,md5buf);
    }
    // SHA256 finalize
    {
        uint64_t bitLen = totalBytes * 8;
        sha256buf[sharem++] = 0x80;
        if (sharem > 56) { memset(sha256buf+sharem,0,64-sharem); SHA256Block(sha,sha256buf); sharem=0; }
        memset(sha256buf+sharem,0,56-sharem);
        // big-endian bit length
        for (int i=0;i<8;i++) sha256buf[63-i]=(uint8_t)(bitLen>>(8*i));
        SHA256Block(sha,sha256buf);
    }

    wchar_t out[200];
    StringCchPrintfW(out, 200,
        L"MD5: %08X%08X%08X%08X\nSHA256: %08X%08X%08X%08X%08X%08X%08X%08X",
        _byteswap_ulong(md5s[0]),_byteswap_ulong(md5s[1]),
        _byteswap_ulong(md5s[2]),_byteswap_ulong(md5s[3]),
        sha[0],sha[1],sha[2],sha[3],sha[4],sha[5],sha[6],sha[7]);
    return out;
}

// MD5 helpers
static inline uint32_t rotl32(uint32_t x, int n){ return (x<<n)|(x>>(32-n)); }
static void MD5Block(uint32_t s[4], const uint8_t b[64]) {
    static const uint32_t K[64]={
        0xd76aa478u,0xe8c7b756u,0x242070dbu,0xc1bdceeeu,0xf57c0fafu,0x4787c62au,0xa8304613u,0xfd469501u,
        0x698098d8u,0x8b44f7afu,0xffff5bb1u,0x895cd7beu,0x6b901122u,0xfd987193u,0xa679438eu,0x49b40821u,
        0xf61e2562u,0xc040b340u,0x265e5a51u,0xe9b6c7aau,0xd62f105du,0x02441453u,0xd8a1e681u,0xe7d3fbc8u,
        0x21e1cde6u,0xc33707d6u,0xf4d50d87u,0x455a14edu,0xa9e3e905u,0xfcefa3f8u,0x676f02d9u,0x8d2a4c8au,
        0xfffa3942u,0x8771f681u,0x6d9d6122u,0xfde5380cu,0xa4beea44u,0x4bdecfa9u,0xf6bb4b60u,0xbebfbc70u,
        0x289b7ec6u,0xeaa127fau,0xd4ef3085u,0x04881d05u,0xd9d4d039u,0xe6db99e5u,0x1fa27cf8u,0xc4ac5665u,
        0xf4292244u,0x432aff97u,0xab9423a7u,0xfc93a039u,0x655b59c3u,0x8f0ccc92u,0xffeff47du,0x85845dd1u,
        0x6fa87e4fu,0xfe2ce6e0u,0xa3014314u,0x4e0811a1u,0xf7537e82u,0xbd3af235u,0x2ad7d2bbu,0xeb86d391u};
    static const int S[64]={7,12,17,22,7,12,17,22,7,12,17,22,7,12,17,22,
                             5, 9,14,20,5, 9,14,20,5, 9,14,20,5, 9,14,20,
                             4,11,16,23,4,11,16,23,4,11,16,23,4,11,16,23,
                             6,10,15,21,6,10,15,21,6,10,15,21,6,10,15,21};
    uint32_t M[16];
    for(int i=0;i<16;i++) M[i]=(uint32_t)b[i*4]|((uint32_t)b[i*4+1]<<8)|((uint32_t)b[i*4+2]<<16)|((uint32_t)b[i*4+3]<<24);
    uint32_t a=s[0],bv=s[1],c=s[2],d=s[3];
    for(int i=0;i<64;i++){
        uint32_t F,g2;
        if(i<16){F=(bv&c)|(~bv&d);g2=(uint32_t)i;}
        else if(i<32){F=(d&bv)|(~d&c);g2=(5*(uint32_t)i+1)%16;}
        else if(i<48){F=bv^c^d;g2=(3*(uint32_t)i+5)%16;}
        else{F=c^(bv|~d);g2=(7*(uint32_t)i)%16;}
        uint32_t tmp=d;d=c;c=bv;bv+=rotl32(a+F+K[i]+M[g2],S[i]);a=tmp;
    }
    s[0]+=a;s[1]+=bv;s[2]+=c;s[3]+=d;
}

static inline uint32_t rotr32(uint32_t x,int n){return(x>>n)|(x<<(32-n));}
static void SHA256Block(uint32_t s[8], const uint8_t b[64]) {
    static const uint32_t K[64]={
        0x428a2f98u,0x71374491u,0xb5c0fbcfu,0xe9b5dba5u,0x3956c25bu,0x59f111f1u,0x923f82a4u,0xab1c5ed5u,
        0xd807aa98u,0x12835b01u,0x243185beu,0x550c7dc3u,0x72be5d74u,0x80deb1feu,0x9bdc06a7u,0xc19bf174u,
        0xe49b69c1u,0xefbe4786u,0x0fc19dc6u,0x240ca1ccu,0x2de92c6fu,0x4a7484aau,0x5cb0a9dcu,0x76f988dau,
        0x983e5152u,0xa831c66du,0xb00327c8u,0xbf597fc7u,0xc6e00bf3u,0xd5a79147u,0x06ca6351u,0x14292967u,
        0x27b70a85u,0x2e1b2138u,0x4d2c6dfcu,0x53380d13u,0x650a7354u,0x766a0abbu,0x81c2c92eu,0x92722c85u,
        0xa2bfe8a1u,0xa81a664bu,0xc24b8b70u,0xc76c51a3u,0xd192e819u,0xd6990624u,0xf40e3585u,0x106aa070u,
        0x19a4c116u,0x1e376c08u,0x2748774cu,0x34b0bcb5u,0x391c0cb3u,0x4ed8aa4au,0x5b9cca4fu,0x682e6ff3u,
        0x748f82eeu,0x78a5636fu,0x84c87814u,0x8cc70208u,0x90befffau,0xa4506cebu,0xbef9a3f7u,0xc67178f2u};
    uint32_t W[64];
    for(int i=0;i<16;i++) W[i]=((uint32_t)b[i*4]<<24)|((uint32_t)b[i*4+1]<<16)|((uint32_t)b[i*4+2]<<8)|(uint32_t)b[i*4+3];
    for(int i=16;i<64;i++){
        uint32_t s0=rotr32(W[i-15],7)^rotr32(W[i-15],18)^(W[i-15]>>3);
        uint32_t s1=rotr32(W[i-2],17)^rotr32(W[i-2],19)^(W[i-2]>>10);
        W[i]=W[i-16]+s0+W[i-7]+s1;
    }
    uint32_t a=s[0],bv=s[1],c=s[2],d=s[3],e=s[4],f=s[5],gv=s[6],hv=s[7];
    for(int i=0;i<64;i++){
        uint32_t S1=rotr32(e,6)^rotr32(e,11)^rotr32(e,25);
        uint32_t ch=(e&f)^(~e&gv);
        uint32_t temp1=hv+S1+ch+K[i]+W[i];
        uint32_t S0=rotr32(a,2)^rotr32(a,13)^rotr32(a,22);
        uint32_t maj=(a&bv)^(a&c)^(bv&c);
        uint32_t temp2=S0+maj;
        hv=gv;gv=f;f=e;e=d+temp1;d=c;c=bv;bv=a;a=temp1+temp2;
    }
    s[0]+=a;s[1]+=bv;s[2]+=c;s[3]+=d;s[4]+=e;s[5]+=f;s[6]+=gv;s[7]+=hv;
}

static void HashWorker() {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    while (!g.hashShutdown.load()) {
        std::wstring path;
        {
            std::unique_lock<std::mutex> lk(g.hashMu);
            g.hashCv.wait(lk, []{ return g.hashShutdown.load() || !g.hashQueue.empty(); });
            if (g.hashShutdown.load()) break;
            path = g.hashQueue.front();
            g.hashQueue.pop();
        }
        {
            std::lock_guard<std::mutex> lk(g.hashMu);
            if (g.hashCache.count(path)) continue;
        }
        std::wstring result = ComputeHash(path);
        {
            std::lock_guard<std::mutex> lk(g.hashMu);
            g.hashCache[path] = result;
        }
        if (g.hwndPreview) {
            wchar_t* dup = (wchar_t*)CoTaskMemAlloc((path.size()+1)*sizeof(wchar_t));
            if (dup) { StringCchCopyW(dup, path.size()+1, path.c_str()); PostMessageW(g.hwndPreview, WM_HASH_READY, 0, (LPARAM)dup); }
        }
    }
    CoUninitialize();
}

static void StartHashThread() {
    g.hashThread = std::thread(HashWorker);
}

static void StopHashThread() {
    g.hashShutdown.store(true);
    g.hashCv.notify_all();
    if (g.hashThread.joinable()) g.hashThread.join();
}

static void RequestHash(const std::wstring& path) {
    {
        std::lock_guard<std::mutex> lk(g.hashMu);
        if (g.hashCache.count(path)) { if (g.hwndPreview) InvalidateRect(g.hwndPreview,nullptr,FALSE); return; }
        std::queue<std::wstring> empty; std::swap(g.hashQueue, empty);
        g.hashQueue.push(path);
    }
    g.hashCv.notify_one();
}

// =====================================================================
// FOLDER SIZE SCANNER
// =====================================================================

static uint64_t ScanFolderSize(const std::wstring& root) {
    uint64_t total = 0;
    std::wstring pat = root + L"\\*";
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pat.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return 0;
    do {
        if (wcscmp(fd.cFileName,L".")==0 || wcscmp(fd.cFileName,L"..")==0) continue;
        if (g.sizeCancel.load()) break;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            total += ScanFolderSize(root + L"\\" + fd.cFileName);
        } else {
            uint64_t sz = ((uint64_t)fd.nFileSizeHigh<<32)|fd.nFileSizeLow;
            total += sz;
            g.sizeBytes.fetch_add(sz);
            g.sizeFiles.fetch_add(1);
        }
        if (g.hwndStatus && (g.sizeFiles.load() % 200) == 0)
            PostMessageW(g.hwndStatus, WM_USER+10, 0, 0);
    } while (FindNextFileW(h,&fd));
    FindClose(h);
    return total;
}

// =====================================================================
// SYSTEM ICONS (extension-cached)
// =====================================================================

static int GetIconIndexForFile(const std::wstring& name, bool isDir) {
    if (isDir) {
        if (g.folderIconIdx >= 0) return g.folderIconIdx;
        SHFILEINFOW sfi{};
        SHGetFileInfoW(L"folder", FILE_ATTRIBUTE_DIRECTORY, &sfi, sizeof(sfi),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
        g.folderIconIdx = sfi.iIcon;
        return g.folderIconIdx;
    }
    std::wstring ext = ExtFromName(name);
    {
        std::lock_guard<std::mutex> lk(g.iconMu);
        auto it = g.iconByExt.find(ext);
        if (it != g.iconByExt.end()) return it->second;
    }
    std::wstring probe = ext.empty() ? std::wstring(L"file") : (std::wstring(L"file.") + ext);
    SHFILEINFOW sfi{};
    SHGetFileInfoW(probe.c_str(), FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi),
        SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
    int idx = sfi.iIcon;
    {
        std::lock_guard<std::mutex> lk(g.iconMu);
        g.iconByExt[ext] = idx;
    }
    return idx;
}

// =====================================================================
// SHELL CONTEXT MENU (gives WinRAR/7-Zip/Compress/Extract/Send To/Open With)
// =====================================================================

static void ReleaseShellMenu() {
    if (g.shellCM3) { g.shellCM3->Release(); g.shellCM3 = nullptr; }
    if (g.shellCM2) { g.shellCM2->Release(); g.shellCM2 = nullptr; }
}

static bool ShowShellContextMenuFor(HWND owner, const std::vector<std::wstring>& paths, POINT screenPt) {
    if (paths.empty()) return false;

    std::vector<LPITEMIDLIST> pidls;
    pidls.reserve(paths.size());
    IShellFolder* parentFolder = nullptr;

    for (const auto& p : paths) {
        LPITEMIDLIST full = nullptr;
        if (FAILED(SHParseDisplayName(p.c_str(), nullptr, &full, 0, nullptr)) || !full) {
            for (auto pi : pidls) CoTaskMemFree(pi);
            if (parentFolder) parentFolder->Release();
            return false;
        }
        if (!parentFolder) {
            LPCITEMIDLIST child = nullptr;
            if (FAILED(SHBindToParent(full, IID_PPV_ARGS(&parentFolder), &child))
                    || !parentFolder) {
                CoTaskMemFree(full);
                for (auto pi : pidls) CoTaskMemFree(pi);
                return false;
            }
            // child points into full — clone it before freeing full
            LPITEMIDLIST childCopy = ILClone(child);
            CoTaskMemFree(full);
            if (!childCopy) { parentFolder->Release(); return false; }
            pidls.push_back(childCopy);
        } else {
            LPCITEMIDLIST child = ILFindLastID(full);
            LPITEMIDLIST childCopy = ILClone(child);
            CoTaskMemFree(full);
            if (!childCopy) {
                for (auto pi : pidls) CoTaskMemFree(pi);
                parentFolder->Release();
                return false;
            }
            pidls.push_back(childCopy);
        }
    }

    std::vector<LPCITEMIDLIST> cpidls;
    cpidls.reserve(pidls.size());
    for (auto p : pidls) cpidls.push_back(p);

    IContextMenu* cm = nullptr;
    HRESULT hr = parentFolder->GetUIObjectOf(owner, (UINT)cpidls.size(),
        cpidls.data(), IID_IContextMenu, nullptr, (void**)&cm);
    parentFolder->Release();
    for (auto p : pidls) CoTaskMemFree(p);

    if (FAILED(hr) || !cm) return false;

    ReleaseShellMenu();
    cm->QueryInterface(IID_PPV_ARGS(&g.shellCM3));
    cm->QueryInterface(IID_PPV_ARGS(&g.shellCM2));

    HMENU hMenu = CreatePopupMenu();
    if (!hMenu) { cm->Release(); ReleaseShellMenu(); return false; }

    const UINT idMin = 1;
    const UINT idMax = 0x7FFF;
    const UINT CMF_ASYNCVERBSTATE_FLAG   = 0x00000400;
    const UINT CMF_OPTIMIZEFORINVOKE_FLAG = 0x00000800;
    bool extended = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    UINT qcmFlags = CMF_NORMAL | CMF_ASYNCVERBSTATE_FLAG | CMF_OPTIMIZEFORINVOKE_FLAG;
    if (extended) qcmFlags |= CMF_EXTENDEDVERBS;
    hr = cm->QueryContextMenu(hMenu, 0, idMin, idMax, qcmFlags);
    if (FAILED(hr)) {
        DestroyMenu(hMenu);
        cm->Release();
        ReleaseShellMenu();
        return false;
    }

    UINT cmd = (UINT)TrackPopupMenu(hMenu,
        TPM_RETURNCMD | TPM_RIGHTBUTTON | TPM_LEFTALIGN,
        screenPt.x, screenPt.y, 0, owner, nullptr);

    bool invoked = false;
    if (cmd > 0) {
        CMINVOKECOMMANDINFOEX info{};
        info.cbSize = sizeof(info);
        info.fMask = CMIC_MASK_UNICODE | CMIC_MASK_PTINVOKE;
        info.hwnd = owner;
        info.lpVerb = MAKEINTRESOURCEA(cmd - idMin);
        info.lpVerbW = MAKEINTRESOURCEW(cmd - idMin);
        info.nShow = SW_SHOWNORMAL;
        info.ptInvoke = screenPt;
        cm->InvokeCommand((LPCMINVOKECOMMANDINFO)&info);
        invoked = true;
    }

    DestroyMenu(hMenu);
    cm->Release();
    ReleaseShellMenu();
    (void)CMF_OPTIMIZEFORINVOKE_FLAG;
    return invoked;
}

// =====================================================================
// GDI HELPERS
// =====================================================================

static void FillRectColor(HDC hdc, const RECT& r, COLORREF c) {
    HBRUSH br = nullptr;
    if (c == COL_BG)        br = g.brBg;
    else if (c == COL_BG_ALT) br = g.brBgAlt;
    else if (c == COL_SEL)  br = g.brSel;
    else if (c == COL_HEADER) br = g.brHeader;
    if (br) { FillRect(hdc, &r, br); return; }
    HBRUSH made = CreateSolidBrush(c);
    FillRect(hdc, &r, made);
    DeleteObject(made);
}

static void DrawHLine(HDC hdc, int x1, int x2, int y, COLORREF c) {
    HPEN p = CreatePen(PS_SOLID, 1, c);
    HPEN old = (HPEN)SelectObject(hdc, p);
    MoveToEx(hdc, x1, y, nullptr);
    LineTo(hdc, x2, y);
    SelectObject(hdc, old);
    DeleteObject(p);
}

static void DrawVLine(HDC hdc, int x, int y1, int y2, COLORREF c) {
    HPEN p = CreatePen(PS_SOLID, 1, c);
    HPEN old = (HPEN)SelectObject(hdc, p);
    MoveToEx(hdc, x, y1, nullptr);
    LineTo(hdc, x, y2);
    SelectObject(hdc, old);
    DeleteObject(p);
}

static void DrawBorderRect(HDC hdc, const RECT& r, COLORREF c) {
    HBRUSH br = CreateSolidBrush(c);
    FrameRect(hdc, &r, br);
    DeleteObject(br);
}

static void DrawTextSimple(HDC hdc, const std::wstring& s, const RECT& r, COLORREF c, UINT flags) {
    SetTextColor(hdc, c);
    SetBkMode(hdc, TRANSPARENT);
    RECT rr = r;
    DrawTextW(hdc, s.c_str(), (int)s.size(), &rr, flags);
}

static std::wstring TruncateToWidth(HDC hdc, const std::wstring& s, int maxw) {
    if (MeasureText(hdc, s) <= maxw) return s;
    std::wstring t = s;
    while (!t.empty()) {
        t.pop_back();
        std::wstring x = t + L"...";
        if (MeasureText(hdc, x) <= maxw) return x;
    }
    return L"...";
}

// =====================================================================
// BACKGROUND RECURSIVE SEARCH
// =====================================================================

struct SearchBatch {
    int paneIdx;
    std::wstring query;
    std::vector<FileEntry> hits;
};

static void BgSearchWorker(HWND paneHwnd, int paneIdx, std::wstring root, std::wstring query) {
    std::vector<std::wstring> dirs;
    dirs.push_back(root);
    auto postBatch = [&](std::vector<FileEntry>& v) {
        if (v.empty()) return;
        auto* b = new SearchBatch();
        b->paneIdx = paneIdx;
        b->query   = query;
        b->hits    = std::move(v);
        PostMessageW(paneHwnd, WM_SEARCH_RESULT, 0, (LPARAM)b);
    };
    std::vector<FileEntry> batch;
    while (!dirs.empty() && !g.searchBgStop.load()) {
        std::wstring cur = dirs.back(); dirs.pop_back();
        std::wstring pat = cur + L"\\*";
        WIN32_FIND_DATAW fd{};
        HANDLE h = FindFirstFileExW(pat.c_str(), FindExInfoBasic, &fd,
            FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH);
        if (h == INVALID_HANDLE_VALUE) continue;
        do {
            if (g.searchBgStop.load()) break;
            if (fd.cFileName[0] == L'.' &&
                (fd.cFileName[1] == L'\0' || (fd.cFileName[1] == L'.' && fd.cFileName[2] == L'\0')))
                continue;
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
                dirs.push_back(cur + L"\\" + fd.cFileName);
            // case-insensitive substring match
            std::wstring nm(fd.cFileName);
            std::wstring nmL = nm, qL = query;
            CharLowerW(&nmL[0]); CharLowerW(&qL[0]);
            if (nmL.find(qL) != std::wstring::npos) {
                FileEntry fe{};
                // Store full path as name so recursive hits are navigable
                fe.name = cur + L"\\" + nm;
                fe.attr = fd.dwFileAttributes;
                ULARGE_INTEGER sz{};
                sz.LowPart  = fd.nFileSizeLow;
                sz.HighPart = fd.nFileSizeHigh;
                fe.size = sz.QuadPart;
                fe.modified = fd.ftLastWriteTime;
                batch.push_back(fe);
                if (batch.size() >= 50) postBatch(batch);
            }
        } while (FindNextFileW(h, &fd));
        FindClose(h);
    }
    postBatch(batch);
}

static void StartBgSearch(Pane& p, const std::wstring& query) {
    // Stop any running search
    g.searchBgStop.store(true);
    if (g.searchBgThread.joinable()) g.searchBgThread.join();
    g.searchBgStop.store(false);
    if (p.tabs.empty()) return;
    std::wstring root = p.tabs[p.activeTab].path;
    HWND phwnd = p.hwnd;
    int pidx = p.paneIndex;
    g.searchBgThread = std::thread(BgSearchWorker, phwnd, pidx, root, query);
    g.searchBgThread.detach();
}

// =====================================================================
// FORWARD DECLS
// =====================================================================

static LRESULT CALLBACK MainProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT CALLBACK SidebarProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT CALLBACK PaneProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT CALLBACK PreviewProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT CALLBACK PathEditSubclass(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
static LRESULT CALLBACK RenameEditSubclass(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
static LRESULT CALLBACK SearchEditSubclass(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
static LRESULT CALLBACK ListSubclass(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
static LRESULT CALLBACK HeaderSubclass(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);

// =====================================================================
// OLE DRAG SOURCE — lets files be dragged to external apps
// =====================================================================

struct OrbDropSource : IDropSource {
    LONG ref = 1;
    ULONG   STDMETHODCALLTYPE AddRef()  override { return (ULONG)InterlockedIncrement(&ref); }
    ULONG   STDMETHODCALLTYPE Release() override {
        LONG r = InterlockedDecrement(&ref); if (!r) delete this; return (ULONG)r;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID r, void** p) override {
        if (r == IID_IUnknown || r == IID_IDropSource) { *p = this; AddRef(); return S_OK; }
        *p = nullptr; return E_NOINTERFACE;
    }
    HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL esc, DWORD keys) override {
        if (esc) return DRAGDROP_S_CANCEL;
        if (!(keys & MK_LBUTTON)) return DRAGDROP_S_DROP;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD) override { return DRAGDROP_S_USEDEFAULTCURSORS; }
};

struct OrbDataObject : IDataObject {
    LONG ref = 1;
    std::vector<std::wstring> files;

    ULONG   STDMETHODCALLTYPE AddRef()  override { return (ULONG)InterlockedIncrement(&ref); }
    ULONG   STDMETHODCALLTYPE Release() override {
        LONG r = InterlockedDecrement(&ref); if (!r) delete this; return (ULONG)r;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID r, void** p) override {
        if (r == IID_IUnknown || r == IID_IDataObject) { *p = this; AddRef(); return S_OK; }
        *p = nullptr; return E_NOINTERFACE;
    }
    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* fmt, STGMEDIUM* med) override {
        if (fmt->cfFormat != CF_HDROP || !(fmt->tymed & TYMED_HGLOBAL)) return DV_E_FORMATETC;
        size_t total = sizeof(DROPFILES) + sizeof(wchar_t); // final double-null
        for (auto& f : files) total += (f.size() + 1) * sizeof(wchar_t);
        HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, total);
        if (!hg) return E_OUTOFMEMORY;
        auto* df = (DROPFILES*)GlobalLock(hg);
        df->pFiles = sizeof(DROPFILES);
        df->fWide  = TRUE;
        wchar_t* dst = (wchar_t*)((BYTE*)df + sizeof(DROPFILES));
        for (auto& f : files) { wcscpy(dst, f.c_str()); dst += f.size() + 1; }
        *dst = L'\0';
        GlobalUnlock(hg);
        med->tymed = TYMED_HGLOBAL;
        med->hGlobal = hg;
        med->pUnkForRelease = nullptr;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* fmt) override {
        return (fmt->cfFormat == CF_HDROP && (fmt->tymed & TYMED_HGLOBAL)) ? S_OK : DV_E_FORMATETC;
    }
    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC*, STGMEDIUM*)   override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC*, FORMATETC* o) override {
        o->ptd = nullptr; return DATA_S_SAMEFORMATETC;
    }
    HRESULT STDMETHODCALLTYPE SetData(FORMATETC*, STGMEDIUM*, BOOL) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD, IEnumFORMATETC**) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD)  override { return OLE_E_ADVISENOTSUPPORTED; }
    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }
};

static void NavigateTab(Pane& p, int tabIdx, const std::wstring& path);
static void RefreshPane(Pane& p);
static void ApplySort(Pane& p, int col);
static void SetActivePane(int idx);
static std::vector<std::wstring> GetSelectedPaths(Pane& p);
static void OpenPath(Pane& p, const std::wstring& path);
static void OpenSelected(Pane& p);
static void GoUp(Pane& p);
static void NewTab(Pane& p, const std::wstring& path);
static void CloseTab(Pane& p, int tabIdx);
static void DoCopy(Pane& p, bool cut);
static void DoPaste(Pane& p);
static void DoDelete(Pane& p, bool perm);
static void DoRename(Pane& p);
static void DoSelectAll(Pane& p);
static void DoNewFolder(Pane& p);
static void TogglePin(const std::wstring& path);
static bool IsPinned(const std::wstring& path);
static void UnpinAt(int idx);
static void LayoutMain(HWND hwnd);
static void UpdatePreview(const std::wstring& path);
static void ShowContextMenu(Pane& p, POINT screenPt);
static void ShowShellContextMenuForPane(Pane& p, POINT screenPt);
static bool ShowShellContextMenuFor(HWND owner, const std::vector<std::wstring>& paths, POINT screenPt);
static int GetIconIndexForFile(const std::wstring& name, bool isDir);
static void GoBack(Pane& p);
static void GoForward(Pane& p);
static void StartFolderSizeForSelection(Pane& p);
static void LoadRecent();
static void SaveRecent();
static void AddRecent(const std::wstring& path);
static void StartHashThread();
static void StopHashThread();
static void RequestHash(const std::wstring& path);
static void UpdateStatusBar();
static void ApplySearchFilter(Pane& p);
static void ShowBulkRenameDialog(Pane& p);
static void DrawToolbar(Pane& p, HDC hdc, const RECT& bar);
static void HandleToolbarClick(Pane& p, int x, int y);
static int  ToolbarHitTest(Pane& p, int x, int y);
static int  EffectiveItemCount(Pane& p);
static FileEntry* EffectiveItem(Pane& p, int idx);
// New feature forward decls
static void SaveSettings();
static void LoadSettings();
static void SaveSession();
static void LoadSession();
static void SaveBookmarks();
static void LoadBookmarks();
static void DoUndo();
static void DoSelectByPattern(Pane& p);
static void DoInvertSelection(Pane& p);
static void DoCopyPath(Pane& p, int mode);
static void DoOpenTerminal(Pane& p);
static void DoZipSelected(Pane& p);
static void DoExtractZip(Pane& p, const std::wstring& zipPath, bool chooseFolder);
static void DoExtractRar(Pane& p, const std::wstring& rarPath, bool chooseFolder);
static void DoSplitFile(Pane& p);
static void DoJoinFiles(Pane& p);
static void DoFileDiff(Pane& p);
static void DoFolderDiff(Pane& p);
static void DoFindDuplicates(Pane& p);
static void ShowFileLockInfo(const std::wstring& path);
static void StartGitStatus(const std::wstring& dir);
static void StartWatchThread(int paneIdx);
static void StopWatchThreads();
static void ShowSettingsDialog();
static void ShowClipHistoryDialog(Pane& p);
static void ShowQuickOpen(Pane& p);
static void ShowCmdPalette(Pane& p);
static void ToggleGridView(Pane& p);
static void ToggleHexView(Pane& p);
static void DoConnectServer(Pane& p);
static void SetColorTag(Pane& p);
static std::wstring SimpleInputBox(HWND parent, const wchar_t* title, const wchar_t* prompt, const wchar_t* defVal);
static void DrawMiniHome(HDC dc, int lx, int cy, COLORREF col);
static void DrawHomePage(Pane& p, HDC hdc, RECT content);
static void NavigateHome(Pane& p);

// =====================================================================
// TOOLBAR (Back / Fwd / Up / Search)
// =====================================================================

static COLORREF LerpColor(COLORREF a, COLORREF b, float t) {
    if (t <= 0.f) return a;
    if (t >= 1.f) return b;
    return RGB(
        (int)(GetRValue(a) + (GetRValue(b) - GetRValue(a)) * t),
        (int)(GetGValue(a) + (GetGValue(b) - GetGValue(a)) * t),
        (int)(GetBValue(a) + (GetBValue(b) - GetBValue(a)) * t));
}

// ---- GDI icon helpers ----

static void DrawTriangleIcon(HDC dc, RECT r, int dir, COLORREF col) {
    // dir: 0=left(back)  1=right(fwd)  2=up
    int cx = (r.left + r.right) / 2;
    int cy = (r.top  + r.bottom) / 2;
    const int S = 5;
    POINT pts[3];
    switch (dir) {
        case 0: pts[0]={cx-S,cy}; pts[1]={cx+S,cy-S}; pts[2]={cx+S,cy+S}; break;
        case 1: pts[0]={cx+S,cy}; pts[1]={cx-S,cy-S}; pts[2]={cx-S,cy+S}; break;
        default:pts[0]={cx,cy-S}; pts[1]={cx+S,cy+S}; pts[2]={cx-S,cy+S}; break;
    }
    HBRUSH br = CreateSolidBrush(col);
    HPEN   np = (HPEN)GetStockObject(NULL_PEN);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)SelectObject(dc, np);
    Polygon(dc, pts, 3);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br);
}

static void DrawGearIcon(HDC dc, RECT r, COLORREF col, COLORREF bgCol) {
    int cx = (r.left + r.right) / 2;
    int cy = (r.top  + r.bottom) / 2;
    const int N = 16; // 8 teeth × 2 alternating radii
    POINT pts[N];
    for (int i = 0; i < N; i++) {
        double angle = i * (2.0 * 3.14159265358979 / N) - 3.14159265358979 / N;
        double rad = (i % 2 == 0) ? 7.0 : 5.0;
        pts[i] = { cx + (int)std::round(rad * std::cos(angle)),
                   cy + (int)std::round(rad * std::sin(angle)) };
    }
    HBRUSH br = CreateSolidBrush(col);
    HPEN   np = (HPEN)GetStockObject(NULL_PEN);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)SelectObject(dc, np);
    Polygon(dc, pts, N);
    // Punch center hole
    HBRUSH hb = CreateSolidBrush(bgCol);
    SelectObject(dc, hb);
    Ellipse(dc, cx - 3, cy - 3, cx + 3, cy + 3);
    DeleteObject(hb);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br);
}

// ---- Sidebar mini-icons ----

static void DrawMiniFolder(HDC dc, int lx, int cy, COLORREF col) {
    HBRUSH br = CreateSolidBrush(col);
    HPEN   np = (HPEN)GetStockObject(NULL_PEN);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)SelectObject(dc, np);
    // Tab nub
    POINT tab[4] = {{lx,cy-1},{lx+4,cy-1},{lx+5,cy-3},{lx,cy-3}};
    Polygon(dc, tab, 4);
    // Body
    RECT body = {lx, cy-1, lx+12, cy+5};
    FillRect(dc, &body, br);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br);
}

static void DrawMiniClock(HDC dc, int lx, int cy, COLORREF col) {
    HBRUSH br  = CreateSolidBrush(col);
    HPEN   pen = CreatePen(PS_SOLID, 1, col);
    HPEN   wp  = CreatePen(PS_SOLID, 1, RGB(255,255,255));
    HPEN   np  = (HPEN)GetStockObject(NULL_PEN);
    HBRUSH ob  = (HBRUSH)SelectObject(dc, br);
    HPEN   op  = (HPEN)SelectObject(dc, np);
    int r = 5, ccx = lx + r, ccy = cy;
    Ellipse(dc, ccx-r, ccy-r, ccx+r+1, ccy+r+1);
    // Hands
    SelectObject(dc, wp);
    MoveToEx(dc, ccx, ccy, nullptr); LineTo(dc, ccx,   ccy-3); // hour
    MoveToEx(dc, ccx, ccy, nullptr); LineTo(dc, ccx+3, ccy+1); // minute
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br); DeleteObject(pen); DeleteObject(wp);
}

static void DrawMiniDrive(HDC dc, int lx, int cy, COLORREF col) {
    HBRUSH br = CreateSolidBrush(col);
    HPEN   np = (HPEN)GetStockObject(NULL_PEN);
    HPEN   lp = CreatePen(PS_SOLID, 1, RGB(200,200,200));
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)SelectObject(dc, np);
    RoundRect(dc, lx, cy-4, lx+12, cy+5, 2, 2);
    SelectObject(dc, lp);
    MoveToEx(dc, lx+2, cy,   nullptr); LineTo(dc, lx+10, cy);
    MoveToEx(dc, lx+2, cy+2, nullptr); LineTo(dc, lx+10, cy+2);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br); DeleteObject(lp);
}

static void DrawMiniCloud(HDC dc, int lx, int cy, COLORREF col) {
    HBRUSH br = CreateSolidBrush(col);
    HPEN   np = (HPEN)GetStockObject(NULL_PEN);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)SelectObject(dc, np);
    Ellipse(dc, lx,   cy-3, lx+7,  cy+4); // left bump
    Ellipse(dc, lx+3, cy-5, lx+10, cy+2); // top bump
    Ellipse(dc, lx+6, cy-3, lx+13, cy+4); // right bump
    RECT base = {lx, cy, lx+13, cy+4};
    FillRect(dc, &base, br);               // fill base
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br);
}

static void DrawMiniHome(HDC dc, int lx, int cy, COLORREF col) {
    HBRUSH br = CreateSolidBrush(col);
    HPEN   np = (HPEN)GetStockObject(NULL_PEN);
    HBRUSH ob = (HBRUSH)SelectObject(dc, br);
    HPEN   op = (HPEN)SelectObject(dc, np);
    // Roof triangle: peak at (lx+6, cy-5), base at (lx-1, cy+1)...(lx+13, cy+1)
    POINT roof[3] = {{lx-1, cy+1}, {lx+6, cy-5}, {lx+13, cy+1}};
    Polygon(dc, roof, 3);
    // Body rectangle below roof
    RECT body = {lx+1, cy+1, lx+11, cy+7};
    FillRect(dc, &body, br);
    // Door cutout (darker rect to simulate door)
    COLORREF doorCol = RGB(
        std::max(0, (int)GetRValue(col) - 50),
        std::max(0, (int)GetGValue(col) - 50),
        std::max(0, (int)GetBValue(col) - 50));
    HBRUSH db = CreateSolidBrush(doorCol);
    RECT door = {lx+4, cy+3, lx+8, cy+7};
    FillRect(dc, &door, db);
    DeleteObject(db);
    SelectObject(dc, ob); SelectObject(dc, op);
    DeleteObject(br);
}

static void DrawHomePage(Pane& p, HDC hdc, RECT content) {
    p.homeItems.clear();
    FillRectColor(hdc, content, COL_BG);

    int pad = 24;
    int y   = content.top + pad;
    int W   = content.right - content.left;
    int availW = W - 2 * pad;

    SelectObject(hdc, g.hFont);

    auto sectionHeader = [&](const wchar_t* title) {
        RECT hr = { content.left + pad, y, content.right - pad, y + 20 };
        SelectObject(hdc, g.hFontBold);
        DrawTextSimple(hdc, title, hr, COL_ACCENT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        SelectObject(hdc, g.hFont);
        DrawHLine(hdc, content.left + pad, content.right - pad, y + 23, COL_BORDER);
        y += 30;
    };

    // ── QUICK ACCESS ──────────────────────────────────────────────────
    sectionHeader(L"Quick Access");

    int tileW = 160, tileH = 70, tileGap = 10;
    int tilesPerRow = std::max(1, (availW + tileGap) / (tileW + tileGap));

    struct QAItem { std::wstring path; std::wstring name; };
    std::vector<QAItem> qa;

    auto tryAdd = [&](const GUID& id, const wchar_t* fallback) {
        std::wstring fp = KnownFolderById(id);
        if (!fp.empty() && DirExists(fp)) {
            std::wstring nm = NameFromPath(fp);
            if (nm.empty()) nm = fallback;
            qa.push_back({fp, nm});
        }
    };
    tryAdd(KFID_Desktop,   L"Desktop");
    tryAdd(KFID_Downloads, L"Downloads");
    tryAdd(KFID_Documents, L"Documents");
    tryAdd(KFID_Pictures,  L"Pictures");
    tryAdd(KFID_Music,     L"Music");
    tryAdd(KFID_Videos,    L"Videos");

    for (auto& pin : g.pinned) {
        bool dup = false;
        for (auto& q : qa) if (q.path == pin) { dup = true; break; }
        if (!dup && DirExists(pin))
            qa.push_back({pin, NameFromPath(pin)});
    }

    int col = 0;
    for (int qi = 0; qi < (int)qa.size(); qi++) {
        int tx = content.left + pad + col * (tileW + tileGap);
        RECT tr = { tx, y, tx + tileW, y + tileH };

        bool hot = (p.homeHover == (int)p.homeItems.size());
        COLORREF bgc = hot ? COL_ACTIVE : COL_BG_ALT;
        HBRUSH tb = CreateSolidBrush(bgc);
        FillRect(hdc, &tr, tb);
        DeleteObject(tb);

        HPEN bp = CreatePen(PS_SOLID, 1, hot ? COL_ACCENT : COL_BORDER);
        HPEN op = (HPEN)SelectObject(hdc, bp);
        HBRUSH nb = (HBRUSH)GetStockObject(NULL_BRUSH);
        HBRUSH ob = (HBRUSH)SelectObject(hdc, nb);
        Rectangle(hdc, tr.left, tr.top, tr.right, tr.bottom);
        SelectObject(hdc, op); DeleteObject(bp);
        SelectObject(hdc, ob);

        // folder icon, centered vertically left side
        DrawMiniFolder(hdc, tr.left + 12, tr.top + tileH / 2 - 2, COL_ACCENT);

        // Name
        RECT nmR = { tr.left + 28, tr.top + 8, tr.right - 6, tr.top + 28 };
        std::wstring nm = TruncateToWidth(hdc, qa[qi].name, nmR.right - nmR.left);
        DrawTextSimple(hdc, nm, nmR, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        // Path (dimmed, smaller)
        RECT prR = { tr.left + 28, tr.top + 30, tr.right - 6, tr.top + tileH - 6 };
        std::wstring ps2 = TruncateToWidth(hdc, qa[qi].path, prR.right - prR.left);
        DrawTextSimple(hdc, ps2, prR, COL_DIM, DT_LEFT | DT_TOP | DT_SINGLELINE);

        p.homeItems.push_back({tr, qa[qi].path});

        col++;
        if (col >= tilesPerRow) { col = 0; y += tileH + tileGap; }
    }
    if (col > 0) y += tileH + tileGap;
    y += 12;

    // ── RECENT ───────────────────────────────────────────────────────
    int maxRecent = std::min((int)g.recent.size(), 12);
    if (maxRecent > 0) {
        sectionHeader(L"Recent");

        int rowH = 30;
        for (int i = 0; i < maxRecent; i++) {
            const std::wstring& rp = g.recent[i];
            RECT rr = { content.left + pad, y, content.right - pad, y + rowH };

            bool hot = (p.homeHover == (int)p.homeItems.size());
            if (hot) {
                HBRUSH hb = CreateSolidBrush(COL_HOVER);
                FillRect(hdc, &rr, hb);
                DeleteObject(hb);
            }

            DrawMiniClock(hdc, rr.left + 8, rr.top + rowH / 2, COL_DIM);

            std::wstring nm = NameFromPath(rp);
            RECT nmR = { rr.left + 24, rr.top, rr.left + 260, rr.bottom };
            nm = TruncateToWidth(hdc, nm, nmR.right - nmR.left);
            DrawTextSimple(hdc, nm, nmR, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

            std::wstring dir = ParentPath(rp);
            RECT dirR = { rr.left + 270, rr.top, rr.right - 4, rr.bottom };
            std::wstring ds = TruncateToWidth(hdc, dir, dirR.right - dirR.left);
            DrawTextSimple(hdc, ds, dirR, COL_DIM, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

            DrawHLine(hdc, rr.left, rr.right, rr.bottom, COL_BORDER);
            p.homeItems.push_back({rr, rp});
            y += rowH;
        }
        y += 12;
    }

    // ── THIS PC ───────────────────────────────────────────────────────
    if (!g.drives.empty() || !g.oneDrivePath.empty()) {
        sectionHeader(L"This PC");

        int dW = 190, dH = 64, dGap = 10;
        int dPerRow = std::max(1, (availW + dGap) / (dW + dGap));
        int dc2 = 0;

        auto drawDrive = [&](const std::wstring& path, const std::wstring& label, bool isCloud) {
            int dx = content.left + pad + dc2 * (dW + dGap);
            RECT dr = { dx, y, dx + dW, y + dH };

            bool hot = (p.homeHover == (int)p.homeItems.size());
            HBRUSH db2 = CreateSolidBrush(hot ? COL_ACTIVE : COL_BG_ALT);
            FillRect(hdc, &dr, db2);
            DeleteObject(db2);

            HPEN bp2 = CreatePen(PS_SOLID, 1, hot ? COL_ACCENT : COL_BORDER);
            HPEN op2 = (HPEN)SelectObject(hdc, bp2);
            HBRUSH nb2 = (HBRUSH)GetStockObject(NULL_BRUSH);
            HBRUSH ob2 = (HBRUSH)SelectObject(hdc, nb2);
            Rectangle(hdc, dr.left, dr.top, dr.right, dr.bottom);
            SelectObject(hdc, op2); DeleteObject(bp2);
            SelectObject(hdc, ob2);

            if (isCloud) DrawMiniCloud(hdc, dr.left + 8, dr.top + 16, COL_ACCENT);
            else         DrawMiniDrive(hdc, dr.left + 8, dr.top + 16, COL_ACCENT);

            RECT lblR = { dr.left + 24, dr.top + 4, dr.right - 4, dr.top + 24 };
            DrawTextSimple(hdc, label, lblR, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

            // Disk space bar
            ULARGE_INTEGER freeB{}, totalB{}, dummy{};
            if (GetDiskFreeSpaceExW(path.c_str(), &freeB, &totalB, &dummy) && totalB.QuadPart > 0) {
                double used = 1.0 - (double)freeB.QuadPart / totalB.QuadPart;
                int bL = dr.left + 8, bR = dr.right - 8;
                int bT = dr.bottom - 12, bB = dr.bottom - 5;
                HBRUSH bgb = CreateSolidBrush(COL_BORDER);
                RECT bgR = {bL, bT, bR, bB};
                FillRect(hdc, &bgR, bgb); DeleteObject(bgb);
                int usedW = (int)((bR - bL) * used);
                COLORREF bc = (used > 0.9) ? RGB(0xC0,0x30,0x30) : COL_ACCENT;
                HBRUSH fgb = CreateSolidBrush(bc);
                RECT fgR = {bL, bT, bL + usedW, bB};
                FillRect(hdc, &fgR, fgb); DeleteObject(fgb);
                // Size label
                wchar_t szb[64];
                uint64_t fG = freeB.QuadPart  / (1024*1024*1024);
                uint64_t tG = totalB.QuadPart / (1024*1024*1024);
                if (tG > 0) StringCchPrintfW(szb, 64, L"%llu GB free of %llu GB", fG, tG);
                else {
                    uint64_t fM = freeB.QuadPart  / (1024*1024);
                    uint64_t tM = totalB.QuadPart / (1024*1024);
                    StringCchPrintfW(szb, 64, L"%llu MB free of %llu MB", fM, tM);
                }
                RECT szR = {bL, bT - 16, bR, bT};
                DrawTextSimple(hdc, szb, szR, COL_DIM, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            }

            p.homeItems.push_back({dr, path});

            dc2++;
            if (dc2 >= dPerRow) { dc2 = 0; y += dH + dGap; }
        };

        if (!g.oneDrivePath.empty()) drawDrive(g.oneDrivePath, L"OneDrive", true);
        for (auto& drv : g.drives) {
            std::wstring lbl = drv;
            if (!lbl.empty() && lbl.back() == L'\\') lbl.pop_back();
            drawDrive(drv, lbl, false);
        }
        if (dc2 > 0) y += dH + dGap;
    }
}

static void NavigateHome(Pane& p) {
    if (p.tabs.empty()) return;
    PaneTab& t = p.tabs[p.activeTab];
    if (t.path == L"::home::") { InvalidateRect(p.hwnd, nullptr, TRUE); return; }
    t.history.resize(t.histPos + 1 > 0 ? t.histPos + 1 : 0);
    t.history.push_back(L"::home::");
    t.histPos = (int)t.history.size() - 1;
    t.path = L"::home::";
    t.search.clear();
    t.filteredEntries.clear();
    g.searchBgStop.store(true);
    if (p.hwndSearch) SetWindowTextW(p.hwndSearch, L"");
    if (p.hwndList) ShowWindow(p.hwndList, SW_HIDE);
    p.isLoading = false;
    KillTimer(p.hwnd, TIMER_SPIN + p.paneIndex);
    p.crumbCached.clear();
    UpdateStatusBar();
    if (g.hwndSidebar) InvalidateRect(g.hwndSidebar, nullptr, FALSE);
    InvalidateRect(p.hwnd, nullptr, TRUE);
}

static void RebuildToolbar(Pane& p, const RECT& bar) {
    p.tbButtons.clear();
    bool canBack = false, canFwd = false;
    if (!p.tabs.empty()) {
        auto& t = p.tabs[p.activeTab];
        canBack = t.histPos > 0;
        canFwd  = t.histPos < (int)t.history.size() - 1;
    }
    int x = bar.left + 4;
    int y = bar.top + 3;
    int bh = bar.bottom - bar.top - 6;
    auto addBtn = [&](const wchar_t* glyph, int id, bool en) {
        ToolbarBtn b;
        b.rect = {x, y, x + TOOLBAR_BTN_W, y + bh};
        b.glyph = glyph; b.id = id; b.enabled = en;
        p.tbButtons.push_back(b);
        x += TOOLBAR_BTN_W + 2;
    };
    addBtn(L"\u2039", IDM_OPEN,    canBack); // ‹ back
    addBtn(L"\u203A", IDM_PASTE,   canFwd);  // › forward
    addBtn(L"\u2191", IDM_REFRESH, true);    // ↑ up

    // ⚙ settings button — right-aligned
    ToolbarBtn gear;
    gear.rect = { bar.right - TOOLBAR_BTN_W - 4, y, bar.right - 4, y + bh };
    gear.glyph = L"\u2699"; gear.id = IDM_SETTINGS; gear.enabled = true;
    p.tbButtons.push_back(gear);
}

static void DrawToolbar(Pane& p, HDC hdc, const RECT& bar) {
    FillRectColor(hdc, bar, COL_BG_ALT);
    DrawHLine(hdc, bar.left, bar.right, bar.bottom - 1, COL_BORDER);
    RebuildToolbar(p, bar);

    HFONT old = (HFONT)SelectObject(hdc, g.hFontBold);
    for (int i = 0; i < (int)p.tbButtons.size(); i++) {
        auto& btn = p.tbButtons[i];
        bool hover   = (p.hoverBtn == i);
        bool active  = (btn.id == IDM_SETTINGS && g.hwndSettings != nullptr);
        float prog   = (i < 8) ? p.tbHoverProg[i] : 0.f;
        COLORREF bg  = active ? COL_ACCENT : LerpColor(COL_BG_ALT, COL_HOVER, prog);
        HBRUSH br = CreateSolidBrush(bg);
        FillRect(hdc, &btn.rect, br);
        DeleteObject(br);
        if (hover || active) DrawBorderRect(hdc, btn.rect, active ? COL_ACCENT : COL_BORDER);
        COLORREF tc = !btn.enabled ? COL_DIM
                    : active       ? RGB(0xFF,0xFF,0xFF)
                    : hover        ? COL_TEXT
                    :                COL_DIM;
        bool isSettings = (btn.id == IDM_SETTINGS);
        if (isSettings)
            DrawGearIcon(hdc, btn.rect, tc, bg);
        else
            DrawTriangleIcon(hdc, btn.rect, i, tc); // i==0 left, 1 right, 2 up
    }
    SelectObject(hdc, old);

    bool isHomeTab = !p.tabs.empty() && p.tabs[p.activeTab].path == L"::home::";
    if (p.hwndSearch) {
        if (isHomeTab) {
            ShowWindow(p.hwndSearch, SW_HIDE);
        } else {
            // leave room for gear button on the right
            int gearW = TOOLBAR_BTN_W + 8;
            int sx = bar.right - SEARCH_W - gearW - 4;
            int sw = SEARCH_W;
            int minSx = bar.left + 3 * (TOOLBAR_BTN_W + 2) + 8;
            if (sx < minSx) sx = minSx;
            MoveWindow(p.hwndSearch, sx, bar.top + 3, sw, bar.bottom - bar.top - 6, FALSE);
            ShowWindow(p.hwndSearch, SW_SHOW);
        }
    }
}

static int ToolbarHitTest(Pane& p, int x, int y) {
    for (int i = 0; i < (int)p.tbButtons.size(); i++)
        if (x >= p.tbButtons[i].rect.left && x < p.tbButtons[i].rect.right &&
            y >= p.tbButtons[i].rect.top  && y < p.tbButtons[i].rect.bottom)
            return i;
    return -1;
}

static void HandleToolbarClick(Pane& p, int x, int y) {
    int idx = ToolbarHitTest(p, x, y);
    if (idx < 0) return;
    auto& btn = p.tbButtons[idx];
    if (!btn.enabled) return;
    if (btn.id == IDM_SETTINGS) {
        if (g.hwndSettings && IsWindow(g.hwndSettings)) {
            DestroyWindow(g.hwndSettings);   // toggle: close if open
        } else {
            ShowSettingsDialog();
        }
        InvalidateRect(p.hwnd, nullptr, FALSE); // redraw gear highlight
        return;
    }
    if (idx == 0) GoBack(p);
    else if (idx == 1) GoForward(p);
    else if (idx == 2) GoUp(p);
}

// =====================================================================
// LISTVIEW HELPERS
// =====================================================================

static std::vector<std::wstring> GetSelectedPaths(Pane& p) {
    std::vector<std::wstring> result;
    if (p.tabs.empty()) return result;
    auto& tab = p.tabs[p.activeTab];
    int n = ListView_GetItemCount(p.hwndList);
    for (int i = 0; i < n; i++) {
        if (ListView_GetItemState(p.hwndList, i, LVIS_SELECTED) & LVIS_SELECTED) {
            FileEntry* e = EffectiveItem(p, i);
            if (e) result.push_back(CombinePath(tab.path, e->name));
        }
    }
    return result;
}

static int GetFocusedIndex(Pane& p) {
    return ListView_GetNextItem(p.hwndList, -1, LVNI_FOCUSED);
}

static void StartFolderSizeForSelection(Pane& p) {
    if (g.sizeRunning.load()) { g.sizeCancel.store(true); return; }
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) return;
    std::wstring root = sel[0];
    DWORD a = GetFileAttributesW(root.c_str());
    if (a == INVALID_FILE_ATTRIBUTES || !(a & FILE_ATTRIBUTE_DIRECTORY)) return;
    g.sizeCancel.store(false);
    g.sizeBytes.store(0);
    g.sizeFiles.store(0);
    g.sizeRunning.store(true);
    g.sizeRoot = NameFromPath(root);
    if (g.sizeThread.joinable()) g.sizeThread.join();
    g.sizeThread = std::thread([root](){
        ScanFolderSize(root);
        g.sizeRunning.store(false);
        if (g.hwndStatus) PostMessageW(g.hwndStatus, WM_USER+10, 0, 0);
    });
}

// =====================================================================
// PANE: NAVIGATION / REFRESH / TABS
// =====================================================================

static void ApplyFolderResult(Pane& p, FolderResult* res) {
    if (p.tabs.empty()) return;
    auto& tab = p.tabs[p.activeTab];
    tab.entries = std::move(res->entries);
    SortEntries(tab.entries, tab.sortCol, tab.sortAsc);
    if (!tab.search.empty()) ApplySearchFilter(p);
    SendMessageW(p.hwndList, WM_SETREDRAW, FALSE, 0);
    ListView_SetItemCountEx(p.hwndList, EffectiveItemCount(p),
        LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
    ListView_SetItemState(p.hwndList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
    int eff = EffectiveItemCount(p);
    if (eff > 0) {
        int f = (tab.focused >= 0 && tab.focused < eff) ? tab.focused : 0;
        ListView_SetItemState(p.hwndList, f, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
        ListView_EnsureVisible(p.hwndList, f, FALSE);
    }
    HDITEMW hd{};
    hd.mask = HDI_FORMAT;
    HWND header = ListView_GetHeader(p.hwndList);
    for (int i = 0; i < 4; i++) {
        Header_GetItem(header, i, &hd);
        hd.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
        if (i == tab.sortCol) hd.fmt |= (tab.sortAsc ? HDF_SORTUP : HDF_SORTDOWN);
        Header_SetItem(header, i, &hd);
    }
    SendMessageW(p.hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(p.hwnd, nullptr, FALSE);
    UpdateWindow(p.hwnd);
    InvalidateRect(p.hwndList, nullptr, TRUE);
    UpdatePreview(L"");
    UpdateStatusBar();

    // Stop spinner
    p.isLoading = false;
    KillTimer(p.hwnd, TIMER_SPIN + p.paneIndex);

    // Fade-in: make list transparent then animate to opaque
    LONG exSt = GetWindowLongW(p.hwndList, GWL_EXSTYLE);
    SetWindowLongW(p.hwndList, GWL_EXSTYLE, exSt | WS_EX_LAYERED);
    p.fadeAlpha = 30;
    SetLayeredWindowAttributes(p.hwndList, 0, 30, LWA_ALPHA);
    SetTimer(p.hwnd, TIMER_FADE + p.paneIndex, 12, nullptr);
}

static void RefreshPane(Pane& p) {
    if (p.tabs.empty()) return;
    auto& tab = p.tabs[p.activeTab];

    if (tab.path == L"::home::") {
        if (p.hwndList) ShowWindow(p.hwndList, SW_HIDE);
        p.isLoading = false;
        KillTimer(p.hwnd, TIMER_SPIN + p.paneIndex);
        p.crumbCached.clear();
        UpdateStatusBar();
        InvalidateRect(p.hwnd, nullptr, TRUE);
        return;
    }

    if (p.hwndList) {
        RECT cr2; GetClientRect(p.hwnd, &cr2);
        int top = TAB_H + CRUMB_H + TOOLBAR_H;
        MoveWindow(p.hwndList, 0, top, cr2.right, cr2.bottom - top, FALSE);
        ShowWindow(p.hwndList, SW_SHOW);
    }

    // Clear list and start spinner
    ListView_SetItemCountEx(p.hwndList, 0, 0);
    if (g.hwndStatus) SetWindowTextW(g.hwndStatus, L"  Loading\u2026");
    p.isLoading = true;
    p.spinFrame = 0;
    SetTimer(p.hwnd, TIMER_SPIN + p.paneIndex, 80, nullptr);

    // Bump serial so any in-flight result from a previous navigation is discarded
    int serial = ++g.paneLoadSerial[p.paneIndex];

    std::wstring path    = tab.path;
    bool showHidden      = g.showHidden;
    HWND paneHwnd        = p.hwnd;
    int  paneIdx         = p.paneIndex;

    std::thread([path, showHidden, paneHwnd, paneIdx, serial]() {
        auto* res = new FolderResult();
        res->paneIdx  = paneIdx;
        res->serial   = serial;
        res->path     = path;
        res->entries  = EnumerateFolder(path, showHidden);
        PostMessageW(paneHwnd, WM_FOLDER_READY, 0, (LPARAM)res);
    }).detach();
}

static void NavigateTabInternal(Pane& p, int tabIdx, const std::wstring& path, bool pushHistory) {
    if (tabIdx < 0 || tabIdx >= (int)p.tabs.size()) return;
    std::wstring n = NormalizePath(path);
    if (n.empty()) return;
    bool goingHome = (n == L"::home::");
    if (!goingHome && !DirExists(n)) return;
    PaneTab& t = p.tabs[tabIdx];
    if (pushHistory && t.path != n) {
        if (t.histPos >= 0 && t.histPos < (int)t.history.size() - 1)
            t.history.resize(t.histPos + 1);
        t.history.push_back(n);
        if (t.history.size() > 200) t.history.erase(t.history.begin());
        t.histPos = (int)t.history.size() - 1;
    }
    t.path = n;
    t.focused = 0;
    t.search.clear();
    t.filteredEntries.clear();
    g.searchBgStop.store(true); // cancel any running recursive search
    if (tabIdx == p.activeTab) {
        if (p.hwndSearch) SetWindowTextW(p.hwndSearch, L"");
        RefreshPane(p);
    }
    InvalidateRect(p.hwnd, nullptr, FALSE);
}

static void NavigateTab(Pane& p, int tabIdx, const std::wstring& path) {
    NavigateTabInternal(p, tabIdx, path, true);
}

static void GoBack(Pane& p) {
    if (p.tabs.empty()) return;
    PaneTab& t = p.tabs[p.activeTab];
    if (t.histPos <= 0) return;
    t.histPos--;
    NavigateTabInternal(p, p.activeTab, t.history[t.histPos], false);
}

static void GoForward(Pane& p) {
    if (p.tabs.empty()) return;
    PaneTab& t = p.tabs[p.activeTab];
    if (t.histPos < 0 || t.histPos >= (int)t.history.size() - 1) return;
    t.histPos++;
    NavigateTabInternal(p, p.activeTab, t.history[t.histPos], false);
}

static int EffectiveItemCount(Pane& p) {
    if (p.tabs.empty()) return 0;
    auto& t = p.tabs[p.activeTab];
    return t.search.empty() ? (int)t.entries.size() : (int)t.filteredEntries.size();
}

static FileEntry* EffectiveItem(Pane& p, int idx) {
    if (p.tabs.empty()) return nullptr;
    auto& t = p.tabs[p.activeTab];
    auto& v = t.search.empty() ? t.entries : t.filteredEntries;
    if (idx < 0 || idx >= (int)v.size()) return nullptr;
    return &v[idx];
}

static void NewTab(Pane& p, const std::wstring& path) {
    std::wstring n = NormalizePath(path);
    if (n.empty() || (n != L"::home::" && !DirExists(n))) return;
    PaneTab t; t.path = n;
    t.history.push_back(n);
    t.histPos = 0;
    p.tabs.push_back(t);
    p.activeTab = (int)p.tabs.size() - 1;
    if (p.hwndSearch) SetWindowTextW(p.hwndSearch, L"");
    RefreshPane(p);
}

static void CloseTab(Pane& p, int tabIdx) {
    if ((int)p.tabs.size() <= 1) return;
    if (tabIdx < 0 || tabIdx >= (int)p.tabs.size()) return;
    p.tabs.erase(p.tabs.begin() + tabIdx);
    if (p.activeTab >= (int)p.tabs.size()) p.activeTab = (int)p.tabs.size() - 1;
    RefreshPane(p);
}

static void GoUp(Pane& p) {
    if (p.tabs.empty()) return;
    auto& t = p.tabs[p.activeTab];
    std::wstring child = NameFromPath(t.path);
    std::wstring parent = ParentPath(t.path);
    if (parent == t.path) return;
    NavigateTab(p, p.activeTab, parent);
    auto& nt = p.tabs[p.activeTab];
    for (size_t i = 0; i < nt.entries.size(); i++) {
        if (CmpStrI(nt.entries[i].name, child) == 0) {
            ListView_SetItemState(p.hwndList, (int)i, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
            ListView_EnsureVisible(p.hwndList, (int)i, FALSE);
            break;
        }
    }
}

static void OpenPath(Pane& p, const std::wstring& path) {
    DWORD a = GetFileAttributesW(path.c_str());
    if (a == INVALID_FILE_ATTRIBUTES) return;
    if (a & FILE_ATTRIBUTE_DIRECTORY) NavigateTab(p, p.activeTab, path);
    else {
        // Check custom associations
        std::wstring ext = ExtFromName(path);
        auto it = g.customAssoc.find(ext);
        if (it != g.customAssoc.end() && !it->second.empty()) {
            ShellExecuteW(nullptr, L"open", it->second.c_str(), (L"\"" + path + L"\"").c_str(), nullptr, SW_SHOWNORMAL);
        } else {
            OpenWithDefault(path);
        }
    }
}

static void OpenSelected(Pane& p) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) {
        int f = GetFocusedIndex(p);
        FileEntry* ep = EffectiveItem(p, f);
        if (!ep) return;
        sel.push_back(CombinePath(p.tabs[p.activeTab].path, ep->name));
    }
    if (sel.size() == 1) {
        DWORD a = GetFileAttributesW(sel[0].c_str());
        if (a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY))
            AddRecent(sel[0]);
        OpenPath(p, sel[0]);
    } else {
        for (auto& s : sel) {
            DWORD a = GetFileAttributesW(s.c_str());
            if (a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY)) {
                AddRecent(s);
                OpenWithDefault(s);
            }
        }
    }
}

static void DoCopy(Pane& p, bool cut) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) return;
    g.clipboardFiles = sel;
    g.clipboardCut = cut;
    g.clipboardCutSet.clear();
    if (cut) for (auto& s : sel) g.clipboardCutSet.insert(s);
    // Track clipboard history
    AppCtx::ClipEntry ce; ce.files = sel; ce.cut = cut;
    g.clipHistory.insert(g.clipHistory.begin(), ce);
    if (g.clipHistory.size() > 8) g.clipHistory.resize(8);
    InvalidateRect(g.panes[0].hwndList, nullptr, FALSE);
    InvalidateRect(g.panes[1].hwndList, nullptr, FALSE);

    if (OpenClipboard(g.hwndMain)) {
        EmptyClipboard();
        auto buf = MakeDoubleNullList(sel);
        size_t bytes = sizeof(DROPFILES) + buf.size() * sizeof(wchar_t);
        HGLOBAL hg = GlobalAlloc(GHND, bytes);
        if (hg) {
            DROPFILES* df = (DROPFILES*)GlobalLock(hg);
            df->pFiles = sizeof(DROPFILES);
            df->fWide = TRUE;
            memcpy((BYTE*)df + sizeof(DROPFILES), buf.data(), buf.size() * sizeof(wchar_t));
            GlobalUnlock(hg);
            SetClipboardData(CF_HDROP, hg);
        }
        CloseClipboard();
    }
}

static std::vector<std::wstring> GetClipboardFiles() {
    std::vector<std::wstring> out;
    if (!OpenClipboard(g.hwndMain)) return out;
    HANDLE h = GetClipboardData(CF_HDROP);
    if (h) {
        HDROP hd = (HDROP)h;
        UINT n = DragQueryFileW(hd, 0xFFFFFFFF, nullptr, 0);
        for (UINT i = 0; i < n; i++) {
            wchar_t b[MAX_PATH];
            if (DragQueryFileW(hd, i, b, MAX_PATH))
                out.push_back(b);
        }
    }
    CloseClipboard();
    return out;
}

static void DoPaste(Pane& p) {
    if (p.tabs.empty()) return;
    auto files = GetClipboardFiles();
    if (files.empty()) files = g.clipboardFiles;
    if (files.empty()) return;
    bool cut = g.clipboardCut;
    if (cut) MoveFiles(g.hwndMain, files, p.tabs[p.activeTab].path);
    else     CopyFiles(g.hwndMain, files, p.tabs[p.activeTab].path);
    if (cut) {
        g.clipboardFiles.clear();
        g.clipboardCutSet.clear();
        g.clipboardCut = false;
    }
    RefreshPane(p);
}

static void DoDelete(Pane& p, bool perm) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) return;
    if (DeleteFiles(g.hwndMain, sel, perm)) RefreshPane(p);
}

static void DoRename(Pane& p) {
    int idx = GetFocusedIndex(p);
    FileEntry* ep = EffectiveItem(p, idx);
    if (!ep) return;
    RECT rc;
    if (!ListView_GetSubItemRect(p.hwndList, idx, 0, LVIR_LABEL, &rc)) return;
    MapWindowPoints(p.hwndList, p.hwnd, (POINT*)&rc, 2);
    rc.right = std::min((LONG)(rc.left + 360), (LONG)(rc.left + 600));
    HWND he = CreateWindowExW(0, L"EDIT", ep->name.c_str(),
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top + 2,
        p.hwnd, (HMENU)(INT_PTR)ID_RENAME_EDIT, g.hInst, nullptr);
    SendMessageW(he, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    SendMessageW(he, EM_SETSEL, 0, -1);
    if (!ep->isDir()) {
        size_t dot = ep->name.find_last_of(L'.');
        if (dot != std::wstring::npos && dot > 0)
            SendMessageW(he, EM_SETSEL, 0, (LPARAM)dot);
    }
    SetWindowSubclass(he, RenameEditSubclass, idx, (DWORD_PTR)&p);
    SetFocus(he);
}

static void DoSelectAll(Pane& p) {
    int n = ListView_GetItemCount(p.hwndList);
    for (int i = 0; i < n; i++)
        ListView_SetItemState(p.hwndList, i, LVIS_SELECTED, LVIS_SELECTED);
}

static void DoNewFolder(Pane& p) {
    if (p.tabs.empty()) return;
    std::wstring nm;
    if (MakeNewFolder(p.tabs[p.activeTab].path, nm)) {
        RefreshPane(p);
        for (size_t i = 0; i < p.tabs[p.activeTab].entries.size(); i++) {
            if (p.tabs[p.activeTab].entries[i].name == nm) {
                ListView_SetItemState(p.hwndList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
                ListView_SetItemState(p.hwndList, (int)i, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
                ListView_EnsureVisible(p.hwndList, (int)i, FALSE);
                DoRename(p);
                break;
            }
        }
    }
}

static void ApplySort(Pane& p, int col) {
    if (p.tabs.empty()) return;
    auto& t = p.tabs[p.activeTab];
    if (t.sortCol == col) t.sortAsc = !t.sortAsc;
    else { t.sortCol = col; t.sortAsc = true; }
    RefreshPane(p);
}

// =====================================================================
// PINS
// =====================================================================

static bool IsPinned(const std::wstring& path) {
    for (auto& p : g.pinned) if (CmpStrI(p, path) == 0) return true;
    return false;
}

static void TogglePin(const std::wstring& path) {
    if (IsPinned(path)) {
        g.pinned.erase(std::remove_if(g.pinned.begin(), g.pinned.end(),
            [&](const std::wstring& s) { return CmpStrI(s, path) == 0; }), g.pinned.end());
    } else {
        g.pinned.push_back(path);
    }
    SavePins();
    InvalidateRect(g.hwndSidebar, nullptr, TRUE);
}

static void UnpinAt(int idx) {
    if (idx < 0 || idx >= (int)g.pinned.size()) return;
    g.pinned.erase(g.pinned.begin() + idx);
    SavePins();
    InvalidateRect(g.hwndSidebar, nullptr, TRUE);
}

// =====================================================================
// SIDEBAR WINDOW
// =====================================================================

static void DrawSidebar(HWND hwnd, HDC hdc) {
    RECT cr; GetClientRect(hwnd, &cr);
    HBRUSH sbBr = CreateSolidBrush(COL_SIDEBAR);
    FillRect(hdc, &cr, sbBr);
    DeleteObject(sbBr);

    auto drawSection = [&](const wchar_t* title, int y) {
        RECT hr = { 10, y + 2, cr.right - 6, y + SIDEBAR_HEADER_H };
        HFONT old2 = (HFONT)SelectObject(hdc, g.hFontBold);
        DrawTextSimple(hdc, title, hr, COL_ACCENT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        SelectObject(hdc, old2);
    };

    // Animated hover fill — lerps between sidebar bg and hover color
    auto drawRowHover = [&](const RECT& r, int rowIdx) {
        float prog = 0.f;
        if (rowIdx == g.sidebarHover)     prog = g.sidebarHoverProg;
        if (rowIdx == g.sidebarPrevHover) prog = std::max(prog, g.sidebarPrevProg);
        if (prog > 0.001f) {
            COLORREF c = LerpColor(COL_SIDEBAR, COL_HOVER, prog);
            HBRUSH hb = CreateSolidBrush(c);
            FillRect(hdc, &r, hb);
            DeleteObject(hb);
        }
    };

    g.sidebarRects.clear();
    SelectObject(hdc, g.hFont);

    int y = 4;

    // Home button (always index 0)
    {
        g.sidebarHomeIdx = (int)g.sidebarRects.size();
        RECT r = { 0, y, cr.right, y + SIDEBAR_ROW_H };
        g.sidebarRects.push_back(r);
        drawRowHover(r, g.sidebarHomeIdx);
        Pane& ap2 = g.panes[g.activePane];
        bool apHome = !ap2.tabs.empty() && ap2.tabs[ap2.activeTab].path == L"::home::";
        if (apHome || g.sidebarHomeIdx == g.sidebarHover) {
            HPEN hpen = CreatePen(PS_SOLID, 2, COL_ACCENT);
            HPEN opn = (HPEN)SelectObject(hdc, hpen);
            MoveToEx(hdc, 0, r.top + 2, nullptr);
            LineTo(hdc, 0, r.bottom - 2);
            SelectObject(hdc, opn); DeleteObject(hpen);
        }
        DrawMiniHome(hdc, 5, y + SIDEBAR_ROW_H / 2 - 1, apHome ? COL_ACCENT : COL_DIM);
        RECT tr = { 22, y, cr.right - 6, y + SIDEBAR_ROW_H };
        DrawTextSimple(hdc, L"Home", tr, apHome ? COL_TEXT : COL_DIM, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        y += SIDEBAR_ROW_H;
        DrawHLine(hdc, 6, cr.right - 6, y, COL_BORDER);
        y += 6;
    }
    g.sidebarPinnedBase = (int)g.sidebarRects.size();

    drawSection(L"PINNED", y);
    y += SIDEBAR_HEADER_H;

    for (size_t i = 0; i < g.pinned.size(); i++) {
        RECT r = { 0, y, cr.right, y + SIDEBAR_ROW_H };
        g.sidebarRects.push_back(r);
        drawRowHover(r, g.sidebarPinnedBase + (int)i);
        // active-pane indicator bar
        if (g.sidebarPinnedBase + (int)i == g.sidebarHover) {
            HPEN ap = CreatePen(PS_SOLID, 2, COL_ACCENT);
            HPEN op = (HPEN)SelectObject(hdc, ap);
            MoveToEx(hdc, 0, r.top + 2, nullptr);
            LineTo(hdc, 0, r.bottom - 2);
            SelectObject(hdc, op); DeleteObject(ap);
        }
        std::wstring nm = NameFromPath(g.pinned[i]);
        if (nm.empty()) nm = g.pinned[i];
        DrawMiniFolder(hdc, 5, y + SIDEBAR_ROW_H/2, COL_ACCENT);
        RECT tr = { 22, y, cr.right - 6, y + SIDEBAR_ROW_H };
        nm = TruncateToWidth(hdc, nm, tr.right - tr.left);
        DrawTextSimple(hdc, nm, tr, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        y += SIDEBAR_ROW_H;
    }

    // Recent files section
    int maxRecent = std::min((int)g.recent.size(), 10);
    if (maxRecent > 0) {
        y += 6;
        DrawHLine(hdc, 8, cr.right - 8, y, COL_BORDER);
        y += 6;
        drawSection(L"RECENT", y);
        y += SIDEBAR_HEADER_H;
        int recentBase = (int)g.sidebarRects.size();
        for (int i = 0; i < maxRecent; i++) {
            RECT r = { 0, y, cr.right, y + SIDEBAR_ROW_H };
            g.sidebarRects.push_back(r);
            drawRowHover(r, recentBase + i);
            std::wstring nm = NameFromPath(g.recent[i]);
            DrawMiniClock(hdc, 5, y + SIDEBAR_ROW_H/2, COL_DIM);
            RECT tr = { 22, y, cr.right - 6, y + SIDEBAR_ROW_H };
            nm = TruncateToWidth(hdc, nm, tr.right - tr.left);
            DrawTextSimple(hdc, nm, tr, COL_DIM, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            y += SIDEBAR_ROW_H;
        }
    }

    // THIS PC section — OneDrive + local drives
    {
        y += 6;
        DrawHLine(hdc, 8, cr.right - 8, y, COL_BORDER);
        y += 6;
        drawSection(L"THIS PC", y);
        y += SIDEBAR_HEADER_H;

        g.sidebarDriveBase = (int)g.sidebarRects.size();

        // OneDrive entry (if detected)
        if (!g.oneDrivePath.empty()) {
            RECT r = { 0, y, cr.right, y + SIDEBAR_ROW_H };
            g.sidebarRects.push_back(r);
            drawRowHover(r, g.sidebarDriveBase + 0);
            DrawMiniCloud(hdc, 5, y + SIDEBAR_ROW_H/2, COL_ACCENT);
            RECT tr = { 22, y, cr.right - 6, y + SIDEBAR_ROW_H };
            DrawTextSimple(hdc, L"OneDrive", tr, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            y += SIDEBAR_ROW_H;
        }

        // Local drive letters
        for (size_t di = 0; di < g.drives.size(); di++) {
            RECT r = { 0, y, cr.right, y + SIDEBAR_ROW_H };
            g.sidebarRects.push_back(r);
            int offset = g.oneDrivePath.empty() ? 0 : 1;
            drawRowHover(r, g.sidebarDriveBase + (int)di + offset);
            // Label: "C:\" style
            std::wstring driveLabel = g.drives[di];
            if (!driveLabel.empty() && driveLabel.back() == L'\\')
                driveLabel.pop_back();
            DrawMiniDrive(hdc, 5, y + SIDEBAR_ROW_H/2, COL_DIM);
            RECT tr = { 22, y, cr.right - 6, y + SIDEBAR_ROW_H };
            DrawTextSimple(hdc, driveLabel, tr, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            y += SIDEBAR_ROW_H;
        }
    }

    DrawVLine(hdc, cr.right - 1, 0, cr.bottom, COL_BORDER);
}

static int SidebarHitTest(HWND hwnd, int x, int y) {
    (void)hwnd;
    for (size_t i = 0; i < g.sidebarRects.size(); i++) {
        const RECT& r = g.sidebarRects[i];
        if (x >= r.left && x < r.right && y >= r.top && y < r.bottom) return (int)i;
    }
    return -1;
}

static LRESULT CALLBACK SidebarProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_CREATE:
            return 0;
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            HDC mem = CreateCompatibleDC(hdc);
            RECT cr; GetClientRect(hwnd, &cr);
            HBITMAP bmp = CreateCompatibleBitmap(hdc, cr.right, cr.bottom);
            HBITMAP oldb = (HBITMAP)SelectObject(mem, bmp);
            SelectObject(mem, g.hFont);
            DrawSidebar(hwnd, mem);
            BitBlt(hdc, 0, 0, cr.right, cr.bottom, mem, 0, 0, SRCCOPY);
            SelectObject(mem, oldb);
            DeleteObject(bmp);
            DeleteDC(mem);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_MOUSEMOVE: {
            int x = GET_X_LPARAM(l), y = GET_Y_LPARAM(l);
            int h = SidebarHitTest(hwnd, x, y);
            if (h != g.sidebarHover) {
                g.sidebarPrevHover = g.sidebarHover;
                g.sidebarPrevProg  = g.sidebarHoverProg;
                g.sidebarHover     = h;
                g.sidebarHoverProg = 0.f;
                SetTimer(hwnd, TIMER_SIDEBAR, 16, nullptr);
                TRACKMOUSEEVENT t{ sizeof(t), TME_LEAVE, hwnd, 0 };
                TrackMouseEvent(&t);
            }
            return 0;
        }
        case WM_MOUSELEAVE: {
            g.sidebarPrevHover = g.sidebarHover;
            g.sidebarPrevProg  = g.sidebarHoverProg;
            g.sidebarHover     = -1;
            g.sidebarHoverProg = 0.f;
            SetTimer(hwnd, TIMER_SIDEBAR, 16, nullptr);
            return 0;
        }
        case WM_TIMER: {
            if (w == TIMER_SIDEBAR) {
                bool settled = true;
                if (g.sidebarHoverProg < 1.f) {
                    g.sidebarHoverProg = std::min(1.f, g.sidebarHoverProg + 0.18f);
                    settled = false;
                }
                if (g.sidebarPrevProg > 0.f) {
                    g.sidebarPrevProg = std::max(0.f, g.sidebarPrevProg - 0.18f);
                    settled = false;
                }
                InvalidateRect(hwnd, nullptr, FALSE);
                if (settled) KillTimer(hwnd, TIMER_SIDEBAR);
            }
            return 0;
        }
        case WM_LBUTTONDOWN: {
            int x = GET_X_LPARAM(l), y = GET_Y_LPARAM(l);
            int h = SidebarHitTest(hwnd, x, y);
            Pane& ap = g.panes[g.activePane];
            if (h == g.sidebarHomeIdx) {
                NavigateHome(ap);
            } else if (h >= g.sidebarPinnedBase && h < g.sidebarPinnedBase + (int)g.pinned.size()) {
                NavigateTab(ap, ap.activeTab, g.pinned[h - g.sidebarPinnedBase]);
                SetFocus(ap.hwndList);
            } else {
                int recentBase = g.sidebarPinnedBase + (int)g.pinned.size();
                int maxRecent = std::min((int)g.recent.size(), 10);
                int recentIdx = h - recentBase;
                if (recentIdx >= 0 && recentIdx < maxRecent) {
                    std::wstring rp = g.recent[recentIdx];
                    DWORD a = GetFileAttributesW(rp.c_str());
                    if (a != INVALID_FILE_ATTRIBUTES) {
                        if (a & FILE_ATTRIBUTE_DIRECTORY)
                            NavigateTab(ap, ap.activeTab, rp);
                        else {
                            std::wstring parent = ParentPath(rp);
                            NavigateTab(ap, ap.activeTab, parent);
                        }
                        SetFocus(ap.hwndList);
                    }
                } else if (h >= g.sidebarDriveBase) {
                    int driveIdx = h - g.sidebarDriveBase;
                    std::wstring target;
                    if (!g.oneDrivePath.empty()) {
                        if (driveIdx == 0) target = g.oneDrivePath;
                        else driveIdx--;
                    }
                    if (target.empty() && driveIdx >= 0 && driveIdx < (int)g.drives.size())
                        target = g.drives[driveIdx];
                    if (!target.empty()) {
                        NavigateTab(ap, ap.activeTab, target);
                        SetFocus(ap.hwndList);
                    }
                }
            }
            return 0;
        }
        case WM_RBUTTONUP: {
            int x = GET_X_LPARAM(l), y = GET_Y_LPARAM(l);
            int h = SidebarHitTest(hwnd, x, y);
            int pidx = h - g.sidebarPinnedBase;
            if (pidx >= 0 && pidx < (int)g.pinned.size()) {
                HMENU m = CreatePopupMenu();
                AppendMenuW(m, MF_STRING, IDM_OPEN, L"Open");
                AppendMenuW(m, MF_STRING, IDM_OPEN_NEW_TAB, L"Open in New Tab");
                AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
                AppendMenuW(m, MF_STRING, IDM_UNPIN_SIDEBAR, L"Remove Pin");
                POINT pt{ x, y };
                ClientToScreen(hwnd, &pt);
                int cmd = TrackPopupMenu(m, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
                DestroyMenu(m);
                Pane& ap = g.panes[g.activePane];
                if (cmd == IDM_OPEN) {
                    NavigateTab(ap, ap.activeTab, g.pinned[pidx]);
                    SetFocus(ap.hwndList);
                } else if (cmd == IDM_OPEN_NEW_TAB) {
                    NewTab(ap, g.pinned[pidx]);
                    SetFocus(ap.hwndList);
                } else if (cmd == IDM_UNPIN_SIDEBAR) {
                    UnpinAt(pidx);
                }
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;
    }
    return DefWindowProc(hwnd, msg, w, l);
}

// =====================================================================
// PREVIEW WINDOW
// =====================================================================

static void DrawPreview(HWND hwnd, HDC hdc) {
    RECT cr; GetClientRect(hwnd, &cr);
    FillRectColor(hdc, cr, COL_BG);
    DrawVLine(hdc, 0, 0, cr.bottom, COL_BORDER);

    SelectObject(hdc, g.hFontBold);
    RECT hr = { 12, 0, cr.right - 8, 26 };
    DrawTextSimple(hdc, L"PREVIEW", hr, COL_DIM, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    DrawHLine(hdc, 0, cr.right, 26, COL_BORDER);
    SelectObject(hdc, g.hFont);

    if (g.previewPath.empty()) {
        RECT msg = { 0, 0, cr.right, cr.bottom };
        DrawTextSimple(hdc, L"No selection", msg, COL_DIM, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        return;
    }

    std::wstring ext = ExtFromName(g.previewPath);
    HBITMAP bmp = nullptr;
    std::wstring meta;
    {
        std::lock_guard<std::mutex> lk(g.thumbMu);
        auto it = g.thumbCache.find(g.previewPath);
        if (it != g.thumbCache.end()) bmp = it->second;
        auto m = g.metaCache.find(g.previewPath);
        if (m != g.metaCache.end()) meta = m->second;
    }

    int top = 32;
    int areaW = cr.right - 16;
    int areaH = (cr.bottom - top) - 140;
    if (areaH < 80) areaH = 80;
    RECT thumbArea = { 8, top, 8 + areaW, top + areaH };

    if (bmp && (IsImageExt(ext) || IsVideoExt(ext) || IsAudioExt(ext))) {
        BITMAP bm{};
        GetObjectW(bmp, sizeof(bm), &bm);
        if (bm.bmWidth > 0 && bm.bmHeight > 0) {
            int dw = areaW - PREVIEW_THUMB_PAD * 2;
            int dh = areaH - PREVIEW_THUMB_PAD * 2;
            double sx = (double)dw / bm.bmWidth;
            double sy = (double)dh / bm.bmHeight;
            double sc = (sx < sy) ? sx : sy;
            if (sc > 1.0) sc = 1.0;
            int rw = (int)(bm.bmWidth * sc);
            int rh = (int)(bm.bmHeight * sc);
            int rx = thumbArea.left + (areaW - rw) / 2;
            int ry = thumbArea.top + (areaH - rh) / 2;
            HDC mem = CreateCompatibleDC(hdc);
            HBITMAP old = (HBITMAP)SelectObject(mem, bmp);
            SetStretchBltMode(hdc, HALFTONE);
            SetBrushOrgEx(hdc, 0, 0, nullptr);
            StretchBlt(hdc, rx, ry, rw, rh, mem, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY);
            SelectObject(mem, old);
            DeleteDC(mem);
            RECT br = { rx - 1, ry - 1, rx + rw + 1, ry + rh + 1 };
            DrawBorderRect(hdc, br, COL_BORDER);
        }
    } else if (IsAudioExt(ext)) {
        DrawTextSimple(hdc, L"[ Audio ]", thumbArea, COL_DIM, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    } else if (IsTextExt(ext)) {
        // Read first ~16 KB and render up to 60 lines in Consolas
        HANDLE hf2 = CreateFileW(g.previewPath.c_str(), GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hf2 != INVALID_HANDLE_VALUE) {
            char raw[16384]; DWORD rd = 0;
            ReadFile(hf2, raw, sizeof(raw) - 1, &rd, nullptr);
            CloseHandle(hf2);
            raw[rd] = '\0';
            // Try UTF-8, fall back to ACP
            int wn = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, raw, (int)rd, nullptr, 0);
            if (wn <= 0) wn = MultiByteToWideChar(CP_ACP, 0, raw, (int)rd, nullptr, 0);
            std::wstring wtext(wn + 1, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, raw, (int)rd, &wtext[0], wn);

            HFONT oldF = (HFONT)SelectObject(hdc, g.hFontMono ? g.hFontMono : g.hFont);
            SetTextColor(hdc, COL_TEXT);
            SetBkColor(hdc, COL_BG_ALT);

            // Fill code area background
            HBRUSH codeBr = CreateSolidBrush(COL_BG_ALT);
            FillRect(hdc, &thumbArea, codeBr);
            DeleteObject(codeBr);
            DrawBorderRect(hdc, thumbArea, COL_BORDER);

            const int LH = 15;
            int ty = thumbArea.top + 4;
            int lines = 0;
            size_t pos = 0;
            while (pos <= wtext.size() && lines < 60 && ty + LH <= thumbArea.bottom - 2) {
                size_t nl = wtext.find(L'\n', pos);
                size_t end = (nl == std::wstring::npos) ? wtext.size() : nl;
                std::wstring ln = wtext.substr(pos, end - pos);
                if (!ln.empty() && ln.back() == L'\r') ln.pop_back();
                RECT lr = { thumbArea.left + 5, ty, thumbArea.right - 5, ty + LH };
                DrawTextW(hdc, ln.c_str(), -1, &lr,
                    DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
                ty += LH; lines++;
                pos = (nl == std::wstring::npos) ? wtext.size() + 1 : nl + 1;
            }
            SelectObject(hdc, oldF);
            SetTextColor(hdc, COL_TEXT);
            SetBkMode(hdc, TRANSPARENT);
        } else {
            DrawTextSimple(hdc, L"Cannot read file", thumbArea, COL_DIM, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
    } else {
        DrawTextSimple(hdc, L"No preview available", thumbArea, COL_DIM, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    int textTop = top + areaH + 8;
    DrawHLine(hdc, 8, cr.right - 8, textTop, COL_BORDER);
    textTop += 8;

    std::wstring name = NameFromPath(g.previewPath);
    HDC hh = hdc;
    SelectObject(hh, g.hFontBold);
    RECT nr = { 12, textTop, cr.right - 12, textTop + 22 };
    DrawTextSimple(hh, TruncateToWidth(hh, name, nr.right - nr.left), nr, COL_TEXT,
        DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hh, g.hFont);
    textTop += 24;

    HANDLE hf = CreateFileW(g.previewPath.c_str(), 0, FILE_SHARE_READ | FILE_SHARE_WRITE,
        nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hf != INVALID_HANDLE_VALUE) {
        LARGE_INTEGER sz{};
        GetFileSizeEx(hf, &sz);
        FILETIME mt{};
        GetFileTime(hf, nullptr, nullptr, &mt);
        wchar_t b1[64], b2[96];
        FormatSize((uint64_t)sz.QuadPart, b1, 64);
        FormatTimeFt(mt, b2, 96);
        std::wstring s1 = std::wstring(L"Size: ") + b1;
        std::wstring s2 = std::wstring(L"Modified: ") + b2;
        RECT lr = { 12, textTop, cr.right - 12, textTop + 20 };
        DrawTextSimple(hh, s1, lr, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        textTop += 20;
        lr.top = textTop; lr.bottom = textTop + 20;
        DrawTextSimple(hh, s2, lr, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        textTop += 22;
        CloseHandle(hf);
    }

    auto drawLines = [&](const std::wstring& text) {
        size_t start = 0;
        while (start < text.size() && textTop < cr.bottom - 8) {
            size_t nl = text.find(L'\n', start);
            std::wstring ln = (nl == std::wstring::npos) ? text.substr(start) : text.substr(start, nl - start);
            RECT lr = { 12, textTop, cr.right - 12, textTop + 20 };
            DrawTextSimple(hh, TruncateToWidth(hh, ln, lr.right - lr.left), lr, COL_TEXT,
                DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            textTop += 20;
            if (nl == std::wstring::npos) break;
            start = nl + 1;
        }
    };

    if (!meta.empty()) { textTop += 4; drawLines(meta); }

    std::wstring hash;
    {
        std::lock_guard<std::mutex> lk(g.hashMu);
        auto it = g.hashCache.find(g.previewPath);
        if (it != g.hashCache.end()) hash = it->second;
    }
    if (!hash.empty() && textTop < cr.bottom - 8) {
        textTop += 6;
        DrawHLine(hh, 8, cr.right - 8, textTop, COL_BORDER);
        textTop += 6;
        drawLines(hash);
    } else if (hash.empty() && textTop < cr.bottom - 8) {
        textTop += 6;
        RECT lr = { 12, textTop, cr.right - 12, textTop + 20 };
        DrawTextSimple(hh, L"Computing hash...", lr, COL_DIM, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    }
}

static LRESULT CALLBACK PreviewProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            HDC mem = CreateCompatibleDC(hdc);
            RECT cr; GetClientRect(hwnd, &cr);
            HBITMAP bmp = CreateCompatibleBitmap(hdc, cr.right, cr.bottom);
            HBITMAP oldb = (HBITMAP)SelectObject(mem, bmp);
            SelectObject(mem, g.hFont);
            DrawPreview(hwnd, mem);
            BitBlt(hdc, 0, 0, cr.right, cr.bottom, mem, 0, 0, SRCCOPY);
            SelectObject(mem, oldb);
            DeleteObject(bmp);
            DeleteDC(mem);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_THUMB_READY:
        case WM_HASH_READY: {
            wchar_t* s = (wchar_t*)l;
            if (s) {
                if (g.previewPath == s) InvalidateRect(hwnd, nullptr, FALSE);
                CoTaskMemFree(s);
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;
    }
    return DefWindowProc(hwnd, msg, w, l);
}

static void UpdatePreview(const std::wstring& path) {
    g.previewPath = path;
    if (!path.empty()) {
        RequestThumb(path);
        DWORD a = GetFileAttributesW(path.c_str());
        if (a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY))
            RequestHash(path);
    }
    if (g.hwndPreview) InvalidateRect(g.hwndPreview, nullptr, FALSE);
}

// =====================================================================
// STATUS BAR
// =====================================================================

static void UpdateStatusBar() {
    Pane& p = g.panes[g.activePane];
    int total = EffectiveItemCount(p);
    int selCount = 0;
    uint64_t selBytes = 0;
    if (p.hwndList) {
        int n = ListView_GetItemCount(p.hwndList);
        for (int i = 0; i < n; i++) {
            if (ListView_GetItemState(p.hwndList, i, LVIS_SELECTED) & LVIS_SELECTED) {
                FileEntry* e = EffectiveItem(p, i);
                if (e) {
                    selCount++;
                    if (!e->isDir()) selBytes += e->size;
                }
            }
        }
    }
    wchar_t buf[256];
    if (g.sizeRunning.load()) {
        wchar_t b1[64];
        FormatSize(g.sizeBytes.load(), b1, 64);
        StringCchPrintfW(buf, 256, L"Scanning %s...  %llu files,  %s",
            g.sizeRoot.c_str(), (unsigned long long)g.sizeFiles.load(), b1);
    } else if (selCount > 0) {
        wchar_t b1[64];
        FormatSize(selBytes, b1, 64);
        StringCchPrintfW(buf, 256, L"%d items   |   %d selected   |   %s",
            total, selCount, b1);
    } else {
        StringCchPrintfW(buf, 256, L"%d items", total);
    }
    g.statusText = buf;
    if (g.hwndStatus) InvalidateRect(g.hwndStatus, nullptr, FALSE);
}

static LRESULT CALLBACK StatusProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_USER + 10:
            UpdateStatusBar();
            return 0;
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            RECT cr; GetClientRect(hwnd, &cr);
            HDC mem = CreateCompatibleDC(hdc);
            HBITMAP bmp = CreateCompatibleBitmap(hdc, cr.right, cr.bottom);
            HBITMAP oldb = (HBITMAP)SelectObject(mem, bmp);
            FillRectColor(mem, cr, COL_BG_ALT);
            DrawHLine(mem, 0, cr.right, 0, COL_BORDER);
            HFONT old = (HFONT)SelectObject(mem, g.hFont);
            RECT tr = { 12, 0, cr.right - 12, cr.bottom };
            DrawTextSimple(mem, g.statusText, tr, COL_TEXT, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            SelectObject(mem, old);
            BitBlt(hdc, 0, 0, cr.right, cr.bottom, mem, 0, 0, SRCCOPY);
            SelectObject(mem, oldb);
            DeleteObject(bmp);
            DeleteDC(mem);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_ERASEBKGND: return 1;
    }
    return DefWindowProcW(hwnd, msg, w, l);
}

// =====================================================================
// PANE: TABS + CRUMB DRAW
// =====================================================================

static void RebuildCrumb(Pane& p, HDC hdc) {
    if (p.tabs.empty()) { p.crumbSegs.clear(); p.crumbPaths.clear(); p.crumbCached.clear(); return; }
    std::wstring path = p.tabs[p.activeTab].path;
    if (p.crumbCached == path && !p.crumbSegs.empty()) return;
    if (path == L"::home::") {
        p.crumbSegs  = {L"Home"};
        p.crumbPaths = {L"::home::"};
        p.crumbCached = path;
        (void)hdc;
        return;
    }
    p.crumbSegs.clear();
    p.crumbPaths.clear();
    auto parts = SplitPath(path);
    std::wstring acc;
    for (size_t i = 0; i < parts.size(); i++) {
        if (i == 0) acc = parts[0] + L"\\";
        else if (i == 1) acc = parts[0] + L"\\" + parts[i];
        else acc = acc + L"\\" + parts[i];
        p.crumbSegs.push_back(parts[i]);
        p.crumbPaths.push_back(acc);
    }
    p.crumbCached = path;
    (void)hdc;
}

static void DrawPaneChrome(Pane& p, HDC hdc, const RECT& cr) {
    FillRectColor(hdc, cr, COL_BG);

    RECT tabBar = { cr.left, cr.top, cr.right, cr.top + TAB_H };
    FillRectColor(hdc, tabBar, COL_HEADER);
    DrawHLine(hdc, cr.left, cr.right, tabBar.bottom - 1, COL_BORDER);

    p.tabRects.clear();
    p.tabCloseRects.clear();
    p.tabRects.resize(p.tabs.size());
    p.tabCloseRects.resize(p.tabs.size());

    SelectObject(hdc, g.hFont);
    int x = cr.left + 2;
    bool paneActive = (p.paneIndex == g.activePane);
    for (size_t i = 0; i < p.tabs.size(); i++) {
        std::wstring lbl = TabLabel(p.tabs[i].path);
        if (lbl.empty()) lbl = L"(empty)";
        int textW = MeasureText(hdc, lbl);
        int tabW = textW + 14 + TAB_X_W + 10;
        if (tabW < 80) tabW = 80;
        if (tabW > 220) tabW = 220;
        RECT tr = { x, cr.top, x + tabW, cr.top + TAB_H };
        p.tabRects[i] = tr;
        bool isActive = ((int)i == p.activeTab);
        bool isHover  = ((int)i == p.hoverTab) && !isActive;

        // background
        COLORREF bg = isActive ? COL_BG : (isHover ? COL_HOVER : COL_HEADER);
        HBRUSH br = CreateSolidBrush(bg);
        FillRect(hdc, &tr, br);
        DeleteObject(br);

        // accent line on top of active tab
        if (isActive) {
            COLORREF accent = paneActive ? COL_ACCENT : COL_DIM;
            HPEN ap = CreatePen(PS_SOLID, 3, accent);
            HPEN op = (HPEN)SelectObject(hdc, ap);
            MoveToEx(hdc, tr.left + 1, tr.top + 1, nullptr);
            LineTo(hdc, tr.right - 1, tr.top + 1);
            SelectObject(hdc, op);
            DeleteObject(ap);
            // side borders
            DrawVLine(hdc, tr.left,      tr.top, tr.bottom, COL_BORDER);
            DrawVLine(hdc, tr.right - 1, tr.top, tr.bottom, COL_BORDER);
        } else {
            // separator between inactive tabs
            DrawVLine(hdc, tr.right - 1, tr.top + 5, tr.bottom - 5, COL_BORDER);
        }

        // label
        RECT tt = { tr.left + 9, tr.top + 1, tr.right - TAB_X_W - 5, tr.bottom };
        std::wstring shown = TruncateToWidth(hdc, lbl, tt.right - tt.left);
        COLORREF tc = isActive ? COL_TEXT : COL_DIM;
        DrawTextSimple(hdc, shown, tt, tc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        // close button
        RECT cx = { tr.right - TAB_X_W - 3, tr.top + 5, tr.right - 5, tr.bottom - 5 };
        p.tabCloseRects[i] = cx;
        bool closeHot = (p.hoverTab == (int)i);
        DrawTextSimple(hdc, L"\u00D7", cx, closeHot ? COL_TEXT : COL_DIM,
            DT_CENTER | DT_VCENTER | DT_SINGLELINE);

        x += tabW;
    }

    // new-tab button
    if (p.tabs.size() < 30) {
        RECT plus = { x + 4, cr.top + 5, x + 22, cr.top + TAB_H - 5 };
        DrawTextSimple(hdc, L"+", plus, COL_DIM, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    RECT crumb = { cr.left, cr.top + TAB_H, cr.right, cr.top + TAB_H + CRUMB_H };
    FillRectColor(hdc, crumb, COL_BG_ALT);
    DrawHLine(hdc, cr.left, cr.right, crumb.bottom - 1, COL_BORDER);

    if (!p.hwndPathEdit || !IsWindowVisible(p.hwndPathEdit)) {
        RebuildCrumb(p, hdc);
        p.crumbRects.clear();
        p.crumbRects.resize(p.crumbSegs.size());
        int cx = cr.left + 8;
        for (size_t i = 0; i < p.crumbSegs.size(); i++) {
            std::wstring s = p.crumbSegs[i];
            bool isLast = (i + 1 == p.crumbSegs.size());
            int w = MeasureText(hdc, s) + 10;
            RECT sr = { cx, crumb.top + 2, cx + w, crumb.bottom - 2 };
            p.crumbRects[i] = sr;
            if ((int)i == p.hoverCrumb && !isLast) FillRectColor(hdc, sr, COL_HOVER);
            COLORREF tc = isLast ? COL_TEXT : COL_DIM;
            DrawTextSimple(hdc, s, sr, tc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            cx += w;
            if (!isLast) {
                RECT sep = { cx, crumb.top + 2, cx + CRUMB_SEP_W, crumb.bottom - 2 };
                DrawTextSimple(hdc, L"\u203A", sep, COL_BORDER, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                cx += CRUMB_SEP_W;
            }
        }
    }

    if (p.paneIndex == 0)
        DrawVLine(hdc, cr.right - 1, cr.top, cr.bottom, COL_BORDER);
}

static void HandleTabClick(Pane& p, int x, int y) {
    for (size_t i = 0; i < p.tabCloseRects.size(); i++) {
        RECT cx = p.tabCloseRects[i];
        if (x >= cx.left && x < cx.right && y >= cx.top && y < cx.bottom) {
            if (p.tabs.size() == 1) {
                if (p.paneIndex == 1) {
                    // closing last tab in secondary pane hides it
                    g.showPane2 = false;
                    g.activePane = 0;
                    SetFocus(g.panes[0].hwndList);
                    LayoutMain(g.hwndMain);
                } else {
                    // reset primary pane to Desktop
                    std::wstring desk = KnownFolderById(KFID_Desktop);
                    if (desk.empty() || !DirExists(desk)) desk = L"C:\\";
                    NavigateTab(p, 0, desk);
                    InvalidateRect(p.hwnd, nullptr, FALSE);
                }
            } else {
                CloseTab(p, (int)i);
                InvalidateRect(p.hwnd, nullptr, FALSE);
            }
            return;
        }
    }
    for (size_t i = 0; i < p.tabRects.size(); i++) {
        RECT tr = p.tabRects[i];
        if (x >= tr.left && x < tr.right && y >= tr.top && y < tr.bottom) {
            if ((int)i != p.activeTab) {
                p.activeTab = (int)i;
                RefreshPane(p);
                InvalidateRect(p.hwnd, nullptr, FALSE);
            }
            return;
        }
    }
    int rightmost = p.tabRects.empty() ? 4 : p.tabRects.back().right;
    if (x >= rightmost && x < rightmost + 26) {
        if (!p.tabs.empty()) NewTab(p, p.tabs[p.activeTab].path);
        return;
    }
}

static void HandleCrumbClick(Pane& p, int x, int y) {
    for (size_t i = 0; i < p.crumbRects.size(); i++) {
        RECT r = p.crumbRects[i];
        if (x >= r.left && x < r.right && y >= r.top && y < r.bottom) {
            std::wstring target = p.crumbPaths[i];
            if (target.size() == 3 && target.back() == L'\\') {
            } else if (target.size() > 3 && target.back() == L'\\') target.pop_back();
            NavigateTab(p, p.activeTab, target);
            return;
        }
    }
    if (!p.tabs.empty()) {
        RECT cr; GetClientRect(p.hwnd, &cr);
        if (x >= 0 && x < cr.right && y >= TAB_H && y < TAB_H + CRUMB_H) {
            if (!p.hwndPathEdit) {
                p.hwndPathEdit = CreateWindowExW(0, L"EDIT", L"",
                    WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
                    8, TAB_H + 2, cr.right - 16, CRUMB_H - 4,
                    p.hwnd, (HMENU)(INT_PTR)ID_PATH_EDIT, g.hInst, nullptr);
                SendMessageW(p.hwndPathEdit, WM_SETFONT, (WPARAM)g.hFont, TRUE);
                SetWindowSubclass(p.hwndPathEdit, PathEditSubclass, 0, (DWORD_PTR)&p);
            }
            SetWindowTextW(p.hwndPathEdit, p.tabs[p.activeTab].path.c_str());
            SendMessageW(p.hwndPathEdit, EM_SETSEL, 0, -1);
            ShowWindow(p.hwndPathEdit, SW_SHOW);
            SetFocus(p.hwndPathEdit);
            InvalidateRect(p.hwnd, nullptr, FALSE);
        }
    }
}

static int PaneTabHitTest(Pane& p, int x, int y) {
    for (size_t i = 0; i < p.tabRects.size(); i++) {
        RECT r = p.tabRects[i];
        if (x >= r.left && x < r.right && y >= r.top && y < r.bottom) return (int)i;
    }
    return -1;
}

static int PaneCrumbHitTest(Pane& p, int x, int y) {
    for (size_t i = 0; i < p.crumbRects.size(); i++) {
        RECT r = p.crumbRects[i];
        if (x >= r.left && x < r.right && y >= r.top && y < r.bottom) return (int)i;
    }
    return -1;
}

// =====================================================================
// PANE WINDOW PROC
// =====================================================================

static void ShowEditPath(Pane& p) {
    RECT cr; GetClientRect(p.hwnd, &cr);
    if (!p.hwndPathEdit) {
        p.hwndPathEdit = CreateWindowExW(0, L"EDIT", L"",
            WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            8, TAB_H + 2, cr.right - 16, CRUMB_H - 4,
            p.hwnd, (HMENU)(INT_PTR)ID_PATH_EDIT, g.hInst, nullptr);
        SendMessageW(p.hwndPathEdit, WM_SETFONT, (WPARAM)g.hFont, TRUE);
        SetWindowSubclass(p.hwndPathEdit, PathEditSubclass, 0, (DWORD_PTR)&p);
    }
    if (!p.tabs.empty())
        SetWindowTextW(p.hwndPathEdit, p.tabs[p.activeTab].path.c_str());
    SendMessageW(p.hwndPathEdit, EM_SETSEL, 0, -1);
    ShowWindow(p.hwndPathEdit, SW_SHOW);
    SetFocus(p.hwndPathEdit);
    InvalidateRect(p.hwnd, nullptr, FALSE);
}

static LRESULT HandleListNotify(Pane& p, NMHDR* nh) {
    switch (nh->code) {
        case LVN_GETDISPINFOW: {
            NMLVDISPINFOW* di = (NMLVDISPINFOW*)nh;
            int idx = di->item.iItem;
            int sub = di->item.iSubItem;
            FileEntry* ep = EffectiveItem(p, idx);
            if (!ep) break;
            FileEntry& e = *ep;
            if (di->item.mask & LVIF_TEXT) {
                static thread_local wchar_t buf[260];
                buf[0] = 0;
                switch (sub) {
                    case 0: StringCchCopyW(buf, 260, e.name.c_str()); break;
                    case 1: if (!e.isDir()) FormatSize(e.size, buf, 260); break;
                    case 2: FormatTimeFt(e.modified, buf, 260); break;
                    case 3: StringCchCopyW(buf, 260, e.type.c_str()); break;
                }
                di->item.pszText = buf;
            }
            if ((di->item.mask & LVIF_IMAGE) && sub == 0) {
                di->item.iImage = GetIconIndexForFile(e.name, e.isDir());
            }
            break;
        }
        case LVN_ITEMCHANGED: {
            NMLISTVIEW* lv = (NMLISTVIEW*)nh;
            if ((lv->uChanged & LVIF_STATE) &&
                ((lv->uOldState ^ lv->uNewState) & (LVIS_SELECTED | LVIS_FOCUSED))) {
                int focIdx = ListView_GetNextItem(p.hwndList, -1, LVNI_FOCUSED);
                FileEntry* ep = EffectiveItem(p, focIdx);
                if (ep) {
                    p.tabs[p.activeTab].focused = focIdx;
                    if (!ep->isDir() && p.paneIndex == g.activePane)
                        UpdatePreview(CombinePath(p.tabs[p.activeTab].path, ep->name));
                    else if (ep->isDir() && p.paneIndex == g.activePane)
                        UpdatePreview(L"");
                }
                UpdateStatusBar();
            }
            break;
        }
        case LVN_COLUMNCLICK: {
            NMLISTVIEW* lv = (NMLISTVIEW*)nh;
            ApplySort(p, lv->iSubItem);
            break;
        }
        case NM_DBLCLK: {
            OpenSelected(p);
            break;
        }
        case NM_RCLICK: {
            POINT pt; GetCursorPos(&pt);
            bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            if (shift) ShowShellContextMenuForPane(p, pt);
            else       ShowContextMenu(p, pt);
            break;
        }
        case NM_CLICK: {
            SetActivePane(p.paneIndex);
            break;
        }
        case LVN_BEGINDRAG: {
            NMLISTVIEW* lv = (NMLISTVIEW*)nh;
            auto sel = GetSelectedPaths(p);
            if (sel.empty()) {
                FileEntry* ep = EffectiveItem(p, lv->iItem);
                if (ep) sel.push_back(CombinePath(p.tabs[p.activeTab].path, ep->name));
            }
            if (sel.empty()) break;

            // OLE drag source so files can be dropped on Explorer, desktop, etc.
            auto* data = new OrbDataObject();
            data->files = sel;
            auto* src  = new OrbDropSource();
            DWORD effect = DROPEFFECT_COPY | DROPEFFECT_MOVE;
            DoDragDrop(data, src, effect, &effect);
            src->Release();
            data->Release();

            // Also record internally so internal pane-to-pane drop still works
            g.dragFiles = sel;
            g.dragSourcePane = p.paneIndex;
            break;
        }
        case LVN_KEYDOWN: {
            NMLVKEYDOWN* kd = (NMLVKEYDOWN*)nh;
            switch (kd->wVKey) {
                case VK_RETURN: OpenSelected(p); break;
                case VK_BACK:   GoUp(p); break;
                case VK_DELETE: DoDelete(p, (GetKeyState(VK_SHIFT) & 0x8000) != 0); break;
                case VK_F2:     DoRename(p); break;
            }
            break;
        }
        case NM_CUSTOMDRAW: {
            if (nh->hwndFrom == p.hwndList) {
                NMLVCUSTOMDRAW* cd = (NMLVCUSTOMDRAW*)nh;
                switch (cd->nmcd.dwDrawStage) {
                    case CDDS_PREPAINT:
                        return CDRF_NOTIFYITEMDRAW;
                    case CDDS_ITEMPREPAINT: {
                        int idx = (int)cd->nmcd.dwItemSpec;
                        bool sel = (ListView_GetItemState(p.hwndList, idx, LVIS_SELECTED) & LVIS_SELECTED) != 0;
                        bool isCut = false;
                        FileEntry* epCheck = EffectiveItem(p, idx);
                        if (g.clipboardCut && !g.clipboardCutSet.empty() && !p.tabs.empty() && epCheck) {
                            auto ppPath = CombinePath(p.tabs[p.activeTab].path, epCheck->name);
                            isCut = g.clipboardCutSet.count(ppPath) > 0;
                        }
                        // Check color tag
                        COLORREF textBk = sel ? COL_SEL : COL_BG;
                        if (epCheck && !p.tabs.empty()) {
                            auto fullPath = CombinePath(p.tabs[p.activeTab].path, epCheck->name);
                            auto tagIt = g.colorTags.find(fullPath);
                            if (tagIt != g.colorTags.end() && !sel) {
                                textBk = tagIt->second;
                            }
                        }
                        // Base text color
                        if (isCut) {
                            cd->clrText = COL_DIM;
                        } else if (!sel && epCheck) {
                            auto git = g.gitStatus.find(epCheck->name);
                            if (git != g.gitStatus.end()) {
                                switch (git->second) {
                                    case L'M': cd->clrText = RGB(220,180, 50); break; // modified — yellow
                                    case L'A': cd->clrText = RGB( 80,200, 80); break; // staged   — green
                                    case L'?': cd->clrText = RGB( 80,200,200); break; // untracked— cyan
                                    default:   cd->clrText = COL_TEXT;
                                }
                            } else { cd->clrText = COL_TEXT; }
                        } else { cd->clrText = COL_TEXT; }
                        cd->clrTextBk = textBk;
                        return CDRF_NEWFONT;
                    }
                }
                return CDRF_DODEFAULT;
            }
            break;
        }
        case NM_SETFOCUS: {
            SetActivePane(p.paneIndex);
            break;
        }
    }
    return 0;
}

static LRESULT CALLBACK PaneProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    Pane* pp = (Pane*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    switch (msg) {
        case WM_NCCREATE: {
            CREATESTRUCTW* cs = (CREATESTRUCTW*)l;
            Pane* np = (Pane*)cs->lpCreateParams;
            np->hwnd = hwnd;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)np);
            return DefWindowProcW(hwnd, msg, w, l);
        }
        case WM_CREATE: {
            Pane& p = *pp;
            p.hwndList = CreateWindowExW(0, WC_LISTVIEWW, L"",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_OWNERDATA |
                LVS_SHOWSELALWAYS | LVS_SHAREIMAGELISTS,
                0, 0, 0, 0, hwnd, (HMENU)(INT_PTR)ID_LIST, g.hInst, nullptr);
            SetWindowTheme(p.hwndList, L"", L"");
            ListView_SetExtendedListViewStyle(p.hwndList,
                LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);
            ListView_SetBkColor(p.hwndList, COL_BG);
            ListView_SetTextBkColor(p.hwndList, COL_BG);
            ListView_SetTextColor(p.hwndList, COL_TEXT);
            SendMessageW(p.hwndList, WM_SETFONT, (WPARAM)g.hFont, TRUE);

            LVCOLUMNW c{};
            c.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
            c.fmt = LVCFMT_LEFT;
            c.pszText = (LPWSTR)L"Name"; c.cx = 280; ListView_InsertColumn(p.hwndList, 0, &c);
            c.fmt = LVCFMT_RIGHT; c.pszText = (LPWSTR)L"Size"; c.cx = 90;
            ListView_InsertColumn(p.hwndList, 1, &c);
            c.fmt = LVCFMT_LEFT; c.pszText = (LPWSTR)L"Date Modified"; c.cx = 140;
            ListView_InsertColumn(p.hwndList, 2, &c);
            c.pszText = (LPWSTR)L"Type"; c.cx = 100;
            ListView_InsertColumn(p.hwndList, 3, &c);

            if (g.sysImageList)
                ListView_SetImageList(p.hwndList, g.sysImageList, LVSIL_SMALL);

            // Search edit
            p.hwndSearch = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL,
                0, 0, SEARCH_W, TOOLBAR_H - 6,
                hwnd, (HMENU)(INT_PTR)ID_SEARCH_EDIT, g.hInst, nullptr);
            SendMessageW(p.hwndSearch, WM_SETFONT, (WPARAM)g.hFont, TRUE);
            SetWindowSubclass(p.hwndSearch, SearchEditSubclass, 0, (DWORD_PTR)&p);

            SetWindowSubclass(p.hwndList, ListSubclass, 0, (DWORD_PTR)&p);
            HWND hHdr = ListView_GetHeader(p.hwndList);
            if (hHdr) {
                SetWindowTheme(hHdr, L"", L"");
                SendMessageW(hHdr, WM_SETFONT, (WPARAM)g.hFontBold, TRUE);
                SetWindowSubclass(hHdr, HeaderSubclass, 0, 0);
            }
            return 0;
        }
        case WM_SIZE: {
            Pane& p = *pp;
            RECT cr; GetClientRect(hwnd, &cr);
            int top = TAB_H + CRUMB_H + TOOLBAR_H;
            bool isHome = !p.tabs.empty() && p.tabs[p.activeTab].path == L"::home::";
            if (isHome) {
                ShowWindow(p.hwndList, SW_HIDE);
                InvalidateRect(hwnd, nullptr, FALSE);
            } else {
                ShowWindow(p.hwndList, SW_SHOW);
                MoveWindow(p.hwndList, 0, top, cr.right, cr.bottom - top, TRUE);
            }
            if (p.hwndPathEdit && IsWindowVisible(p.hwndPathEdit))
                MoveWindow(p.hwndPathEdit, 8, TAB_H + 2, cr.right - 16, CRUMB_H - 4, TRUE);
            return 0;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            HDC mem = CreateCompatibleDC(hdc);
            RECT cr; GetClientRect(hwnd, &cr);
            int chromeH = TAB_H + CRUMB_H + TOOLBAR_H;
            bool isHome = !pp->tabs.empty() && pp->tabs[pp->activeTab].path == L"::home::";
            int bmpH = isHome ? cr.bottom : chromeH;
            HBITMAP bmp = CreateCompatibleBitmap(hdc, cr.right, bmpH);
            HBITMAP oldb = (HBITMAP)SelectObject(mem, bmp);
            SelectObject(mem, g.hFont);
            RECT topRect = { 0, 0, cr.right, TAB_H + CRUMB_H };
            DrawPaneChrome(*pp, mem, topRect);
            RECT tbRect = { 0, TAB_H + CRUMB_H, cr.right, chromeH };
            DrawToolbar(*pp, mem, tbRect);

            // Loading spinner — 8 rotating dots in toolbar top-right
            if (pp->isLoading) {
                int cx = cr.right - 22, cy = TAB_H + CRUMB_H + TOOLBAR_H / 2;
                int frame = pp->spinFrame;
                HPEN np = (HPEN)GetStockObject(NULL_PEN);
                for (int si = 0; si < 8; si++) {
                    double ang = (si - frame + 8) % 8 * (2.0 * 3.14159265 / 8);
                    int dx = (int)(9.0 * std::cos(ang));
                    int dy = (int)(9.0 * std::sin(ang));
                    float bright = 0.15f + 0.85f * ((8 - si) / 8.0f);
                    COLORREF c = LerpColor(COL_BG_ALT, COL_ACCENT, bright);
                    HBRUSH db = CreateSolidBrush(c);
                    HBRUSH ob2 = (HBRUSH)SelectObject(mem, db);
                    HPEN   op2 = (HPEN)SelectObject(mem, np);
                    Ellipse(mem, cx+dx-3, cy+dy-3, cx+dx+3, cy+dy+3);
                    SelectObject(mem, ob2); SelectObject(mem, op2);
                    DeleteObject(db);
                }
            }

            if (isHome) {
                RECT homeContent = { 0, chromeH, cr.right, cr.bottom };
                DrawHomePage(*pp, mem, homeContent);
            }

            BitBlt(hdc, 0, 0, cr.right, bmpH, mem, 0, 0, SRCCOPY);
            SelectObject(mem, oldb);
            DeleteObject(bmp);
            DeleteDC(mem);

            // Active pane border — 2px pulsing accent line along top of pane
            if (g.activePane == pp->paneIndex) {
                float pulse = 0.55f + 0.45f * (float)std::sin(GetTickCount() / 380.0);
                COLORREF bc = LerpColor(COL_BG_ALT, COL_ACCENT, pulse);
                HPEN bp = CreatePen(PS_SOLID, 2, bc);
                HPEN op3 = (HPEN)SelectObject(hdc, bp);
                MoveToEx(hdc, 0, 0, nullptr); LineTo(hdc, cr.right, 0);
                SelectObject(hdc, op3); DeleteObject(bp);
            }

            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_LBUTTONDOWN: {
            int x = GET_X_LPARAM(l), y = GET_Y_LPARAM(l);
            if (y < TAB_H) {
                int t = PaneTabHitTest(*pp, x, y);
                if (t >= 0) {
                    // Begin tab drag (tear-off detection)
                    g.tabDragging  = true;
                    g.tabDragPane  = pp->paneIndex;
                    g.tabDragIdx   = t;
                    g.tabTearOut   = false;
                    POINT cp; GetCursorPos(&cp);
                    g.tabDragStart = cp;
                    SetCapture(hwnd);
                }
                HandleTabClick(*pp, x, y);
            }
            else if (y < TAB_H + CRUMB_H) HandleCrumbClick(*pp, x, y);
            else if (y < TAB_H + CRUMB_H + TOOLBAR_H) HandleToolbarClick(*pp, x, y);
            else {
                // Content area — handle home item clicks
                bool isHomeNow = !pp->tabs.empty() && pp->tabs[pp->activeTab].path == L"::home::";
                if (isHomeNow) {
                    for (auto& item : pp->homeItems) {
                        if (x >= item.rect.left && x < item.rect.right &&
                            y >= item.rect.top  && y < item.rect.bottom) {
                            DWORD attr = GetFileAttributesW(item.path.c_str());
                            if (attr != INVALID_FILE_ATTRIBUTES) {
                                if (attr & FILE_ATTRIBUTE_DIRECTORY)
                                    NavigateTab(*pp, pp->activeTab, item.path);
                                else {
                                    std::wstring par = ParentPath(item.path);
                                    NavigateTab(*pp, pp->activeTab, par);
                                }
                            }
                            break;
                        }
                    }
                }
            }
            SetActivePane(pp->paneIndex);
            return 0;
        }
        case WM_MBUTTONDOWN: {
            int x = GET_X_LPARAM(l), y = GET_Y_LPARAM(l);
            if (y < TAB_H) {
                int t = PaneTabHitTest(*pp, x, y);
                if (t >= 0 && pp->tabs.size() > 1) {
                    CloseTab(*pp, t);
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
            }
            return 0;
        }
        case WM_MOUSEMOVE: {
            int x = GET_X_LPARAM(l), y = GET_Y_LPARAM(l);
            // Tab tear-off drag
            if (g.tabDragging && g.tabDragPane == pp->paneIndex) {
                POINT cp; GetCursorPos(&cp);
                int dx = abs(cp.x - g.tabDragStart.x);
                int dy = abs(cp.y - g.tabDragStart.y);
                if (dx > GetSystemMetrics(SM_CXDRAG) || dy > GetSystemMetrics(SM_CYDRAG)) {
                    // Check if cursor has left the pane window
                    RECT wr; GetWindowRect(hwnd, &wr);
                    bool outside = (cp.x < wr.left || cp.x >= wr.right ||
                                    cp.y < wr.top  || cp.y >= wr.bottom);
                    if (outside && !g.tabTearOut) {
                        g.tabTearOut = true;
                        SetCursor(LoadCursorW(nullptr, IDC_SIZEALL));
                        // Visual indicator: redraw pane tab with dashed outline
                        InvalidateRect(hwnd, nullptr, FALSE);
                    }
                    if (!outside && g.tabTearOut) {
                        g.tabTearOut = false;
                        InvalidateRect(hwnd, nullptr, FALSE);
                    }
                }
                return 0;
            }
            int newTab = -1, newCrumb = -1, newBtn = -1;
            if (y < TAB_H) newTab = PaneTabHitTest(*pp, x, y);
            else if (y < TAB_H + CRUMB_H) newCrumb = PaneCrumbHitTest(*pp, x, y);
            else if (y < TAB_H + CRUMB_H + TOOLBAR_H) {
                RECT cr2; GetClientRect(hwnd, &cr2);
                RECT tbR = {0, TAB_H+CRUMB_H, cr2.right, TAB_H+CRUMB_H+TOOLBAR_H};
                RebuildToolbar(*pp, tbR);
                newBtn = ToolbarHitTest(*pp, x, y);
            }
            if (newTab != pp->hoverTab || newCrumb != pp->hoverCrumb || newBtn != pp->hoverBtn) {
                pp->hoverTab   = newTab;
                pp->hoverCrumb = newCrumb;
                pp->hoverBtn   = newBtn;
                SetTimer(hwnd, TIMER_ANIM + pp->paneIndex, 16, nullptr);
                InvalidateRect(hwnd, nullptr, FALSE);
                TRACKMOUSEEVENT t{ sizeof(t), TME_LEAVE, hwnd, 0 };
                TrackMouseEvent(&t);
            }
            // Home page hover tracking
            {
                bool isHomeM = !pp->tabs.empty() && pp->tabs[pp->activeTab].path == L"::home::";
                if (isHomeM && y >= TAB_H + CRUMB_H + TOOLBAR_H) {
                    int newH = -1;
                    for (int hi = 0; hi < (int)pp->homeItems.size(); hi++) {
                        if (x >= pp->homeItems[hi].rect.left && x < pp->homeItems[hi].rect.right &&
                            y >= pp->homeItems[hi].rect.top  && y < pp->homeItems[hi].rect.bottom) {
                            newH = hi; break;
                        }
                    }
                    if (newH != pp->homeHover) {
                        pp->homeHover = newH;
                        InvalidateRect(hwnd, nullptr, FALSE);
                    }
                } else if (pp->homeHover >= 0) {
                    pp->homeHover = -1;
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
            }
            return 0;
        }
        case WM_LBUTTONUP: {
            if (g.tabDragging && g.tabDragPane == pp->paneIndex) {
                bool tearOut = g.tabTearOut;
                int  dragIdx = g.tabDragIdx;
                g.tabDragging = false;
                g.tabTearOut  = false;
                ReleaseCapture();
                if (tearOut && dragIdx >= 0 && dragIdx < (int)pp->tabs.size()) {
                    // Create a new top-level window for this tab
                    std::wstring path = pp->tabs[dragIdx].path;
                    // Close tab from current pane if more than 1 tab
                    if (pp->tabs.size() > 1) CloseTab(*pp, dragIdx);
                    // Spawn new window
                    HWND newWin = CreateWindowExW(0, CLASS_MAIN, L"OrbitFM",
                        WS_OVERLAPPEDWINDOW,
                        CW_USEDEFAULT, CW_USEDEFAULT, 1100, 700,
                        nullptr, nullptr, g.hInst, nullptr);
                    if (newWin) {
                        // Navigate pane[0] of new window to path
                        // The new window starts fresh; we PostMessage a navigate command
                        // We pass the path via a global temp (simple approach for single-file app)
                        SetPropW(newWin, L"InitPath", (HANDLE)0);
                        // Store path in window title temporarily then parse in WM_CREATE
                        std::wstring title = std::wstring(L"OrbitFM") + L"|" + path;
                        SetWindowTextW(newWin, title.c_str());
                        ShowWindow(newWin, SW_SHOW);
                        UpdateWindow(newWin);
                    }
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
                return 0;
            }
            return 0;
        }
        case WM_MOUSELEAVE: {
            pp->hoverTab   = -1;
            pp->hoverCrumb = -1;
            pp->hoverBtn   = -1;
            pp->homeHover  = -1;
            SetTimer(hwnd, TIMER_ANIM + pp->paneIndex, 16, nullptr);
            InvalidateRect(hwnd, nullptr, FALSE);
            return 0;
        }
        case WM_NOTIFY: {
            NMHDR* nh = (NMHDR*)l;
            if (nh->idFrom == ID_LIST || nh->hwndFrom == pp->hwndList) {
                return HandleListNotify(*pp, nh);
            }
            return 0;
        }
        case WM_FOLDER_SIZE:
            UpdateStatusBar();
            return 0;
        case WM_USER + 10:
            UpdateStatusBar();
            return 0;
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC: {
            HDC hdc = (HDC)w;
            SetBkColor(hdc, COL_BG_ALT);
            SetTextColor(hdc, COL_TEXT);
            return (LRESULT)g.brBgAlt;
        }
        case WM_INITMENUPOPUP:
        case WM_DRAWITEM:
        case WM_MEASUREITEM: {
            if (g.shellCM3) {
                LRESULT res = 0;
                if (SUCCEEDED(g.shellCM3->HandleMenuMsg2(msg, w, l, &res))) return res;
            }
            if (g.shellCM2) {
                g.shellCM2->HandleMenuMsg(msg, w, l);
                return (msg == WM_INITMENUPOPUP) ? 0 : TRUE;
            }
            break;
        }
        case WM_MENUCHAR: {
            if (g.shellCM3) {
                LRESULT res = 0;
                if (SUCCEEDED(g.shellCM3->HandleMenuMsg2(msg, w, l, &res))) return res;
            }
            break;
        }
        case WM_AUTOREFRESH: {
            if (pp) {
                // Only refresh if no rename edit box is visible
                bool renaming = false;
                HWND child = GetWindow(hwnd, GW_CHILD);
                while (child) {
                    wchar_t cls[64] = {};
                    GetClassNameW(child, cls, 63);
                    if (_wcsicmp(cls, L"EDIT") == 0 && IsWindowVisible(child)) {
                        int ctlId = GetDlgCtrlID(child);
                        if (ctlId == ID_RENAME_EDIT) { renaming = true; break; }
                    }
                    child = GetWindow(child, GW_HWNDNEXT);
                }
                if (!renaming) RefreshPane(*pp);
            }
            return 0;
        }
        case WM_COMMAND: {
            int id = LOWORD(w);
            // Handle new feature commands forwarded from child windows/palette
            Pane& cp2 = *pp;
            switch (id) {
                case IDM_SELECT_PATTERN:  DoSelectByPattern(cp2); return 0;
                case IDM_INVERT_SEL:      DoInvertSelection(cp2); return 0;
                case IDM_COPY_PATH_FULL:  DoCopyPath(cp2, 0); return 0;
                case IDM_COPY_PATH_NAME:  DoCopyPath(cp2, 1); return 0;
                case IDM_COPY_PATH_DIR:   DoCopyPath(cp2, 2); return 0;
                case IDM_OPEN_TERMINAL:   DoOpenTerminal(cp2); return 0;
                case IDM_ZIP_SELECTED:    DoZipSelected(cp2); return 0;
                case IDM_UNZIP_HERE: {
                    auto sel2 = GetSelectedPaths(cp2);
                    if (!sel2.empty()) DoExtractZip(cp2, sel2[0], false);
                    return 0;
                }
                case IDM_UNZIP_TO_FOLDER: {
                    auto sel2 = GetSelectedPaths(cp2);
                    if (!sel2.empty()) DoExtractZip(cp2, sel2[0], true);
                    return 0;
                }
                case IDM_UNRAR_HERE: {
                    auto sel2 = GetSelectedPaths(cp2);
                    if (!sel2.empty()) DoExtractRar(cp2, sel2[0], false);
                    return 0;
                }
                case IDM_UNRAR_TO_FOLDER: {
                    auto sel2 = GetSelectedPaths(cp2);
                    if (!sel2.empty()) DoExtractRar(cp2, sel2[0], true);
                    return 0;
                }
                case IDM_SPLIT_FILE:      DoSplitFile(cp2); return 0;
                case IDM_JOIN_FILES:      DoJoinFiles(cp2); return 0;
                case IDM_FILE_DIFF:       DoFileDiff(cp2); return 0;
                case IDM_FOLDER_DIFF:     DoFolderDiff(cp2); return 0;
                case IDM_FIND_DUPES:      DoFindDuplicates(cp2); return 0;
                case IDM_PROC_LOCK: {
                    auto sel2 = GetSelectedPaths(cp2);
                    if (!sel2.empty()) ShowFileLockInfo(sel2[0]);
                    return 0;
                }
                case IDM_GRID_VIEW:       ToggleGridView(cp2); return 0;
                case IDM_HEX_VIEW:        ToggleHexView(cp2); return 0;
                case IDM_UNDO:            DoUndo(); return 0;
                case IDM_SETTINGS:        ShowSettingsDialog(); return 0;
                case IDM_CLIP_HISTORY:    ShowClipHistoryDialog(cp2); return 0;
                case IDM_QUICK_OPEN:      ShowQuickOpen(cp2); return 0;
                case IDM_CMD_PALETTE:     ShowCmdPalette(cp2); return 0;
                case IDM_GIT_STATUS:
                    if (!cp2.tabs.empty()) StartGitStatus(cp2.tabs[cp2.activeTab].path);
                    return 0;
                case IDM_CONNECT_SERVER:  DoConnectServer(cp2); return 0;
                default: break;
            }
            // Also handle old WM_COMMAND cases from original code
            {
                int code2 = HIWORD(w);
                if (id == ID_PATH_EDIT && code2 == EN_KILLFOCUS) {
                    if (pp->hwndPathEdit) ShowWindow(pp->hwndPathEdit, SW_HIDE);
                    InvalidateRect(hwnd, nullptr, FALSE);
                    return 0;
                }
                if (id == ID_SEARCH_EDIT && code2 == EN_CHANGE) {
                    wchar_t buf[512] = {};
                    if (pp->hwndSearch) GetWindowTextW(pp->hwndSearch, buf, 511);
                    if (!pp->tabs.empty()) {
                        pp->tabs[pp->activeTab].search = buf;
                        // Instant filter on current-dir entries
                        ApplySearchFilter(*pp);
                        ListView_SetItemCountEx(pp->hwndList, EffectiveItemCount(*pp),
                            LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
                        InvalidateRect(pp->hwndList, nullptr, TRUE);
                        UpdateStatusBar();
                        // Kick off recursive bg search when query is long enough
                        std::wstring q = buf;
                        if (q.size() >= 2) {
                            StartBgSearch(*pp, q);
                        } else {
                            g.searchBgStop.store(true);
                        }
                    }
                    return 0;
                }
            }
            return 0;
        }
        case WM_ERASEBKGND: return 1;

        case WM_FOLDER_READY: {
            if (!pp) break;
            auto* res = (FolderResult*)l;
            // Discard if the pane navigated again before this result arrived
            if (res->serial == g.paneLoadSerial[pp->paneIndex] &&
                !pp->tabs.empty() &&
                pp->tabs[pp->activeTab].path == res->path) {
                ApplyFolderResult(*pp, res);
            }
            delete res;
            return 0;
        }
        case WM_SEARCH_RESULT: {
            if (!pp) break;
            auto* b = (SearchBatch*)l;
            if (!g.searchBgStop.load() &&
                !pp->tabs.empty() &&
                pp->tabs[pp->activeTab].search == b->query)
            {
                auto& filtered = pp->tabs[pp->activeTab].filteredEntries;
                for (auto& fe : b->hits) filtered.push_back(fe);
                ListView_SetItemCountEx(pp->hwndList, (int)filtered.size(),
                    LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
                UpdateStatusBar();
            }
            delete b;
            return 0;
        }
        case WM_TIMER: {
            if (!pp) break;
            Pane& fp = *pp;

            // Folder fade-in
            if ((int)w == TIMER_FADE + fp.paneIndex) {
                fp.fadeAlpha = std::min(255, fp.fadeAlpha + 50);
                SetLayeredWindowAttributes(fp.hwndList, 0, (BYTE)fp.fadeAlpha, LWA_ALPHA);
                if (fp.fadeAlpha >= 255) {
                    KillTimer(hwnd, w);
                    SetWindowLongW(fp.hwndList, GWL_EXSTYLE,
                        GetWindowLongW(fp.hwndList, GWL_EXSTYLE) & ~WS_EX_LAYERED);
                    InvalidateRect(fp.hwndList, nullptr, FALSE);
                }
                return 0;
            }

            // Toolbar hover glow
            if ((int)w == TIMER_ANIM + fp.paneIndex) {
                bool settled = true;
                int n = (int)fp.tbButtons.size();
                for (int i = 0; i < 8; i++) {
                    float target = (i < n && i == fp.hoverBtn) ? 1.f : 0.f;
                    float& prog  = fp.tbHoverProg[i];
                    float  step  = 0.14f;
                    if (prog < target) { prog = std::min(1.f, prog + step); settled = false; }
                    else if (prog > target) { prog = std::max(0.f, prog - step); settled = false; }
                }
                RECT cr2; GetClientRect(hwnd, &cr2);
                RECT tbR = {0, TAB_H+CRUMB_H, cr2.right, TAB_H+CRUMB_H+TOOLBAR_H};
                InvalidateRect(hwnd, &tbR, FALSE);
                if (settled) KillTimer(hwnd, w);
                return 0;
            }

            // Loading spinner rotation
            if ((int)w == TIMER_SPIN + fp.paneIndex) {
                fp.spinFrame = (fp.spinFrame + 1) % 8;
                RECT cr2; GetClientRect(hwnd, &cr2);
                RECT spinR = {cr2.right - 44, TAB_H+CRUMB_H, cr2.right, TAB_H+CRUMB_H+TOOLBAR_H};
                InvalidateRect(hwnd, &spinR, FALSE);
                return 0;
            }

            // Active pane border pulse
            if ((int)w == TIMER_PULSE + fp.paneIndex) {
                RECT cr2; GetClientRect(hwnd, &cr2);
                RECT topLine = {0, 0, cr2.right, 3};
                InvalidateRect(hwnd, &topLine, FALSE);
                return 0;
            }

            break;
        }
    }
    return DefWindowProcW(hwnd, msg, w, l);
}

// =====================================================================
// LIST SUBCLASS / EDIT SUBCLASSES
// =====================================================================

static LRESULT CALLBACK HeaderSubclass(HWND hwnd, UINT msg, WPARAM w, LPARAM l, UINT_PTR id, DWORD_PTR ref) {
    (void)id; (void)ref;
    if (msg == WM_ERASEBKGND) {
        HDC hdc = (HDC)w;
        RECT cr; GetClientRect(hwnd, &cr);
        FillRectColor(hdc, cr, COL_HEADER);
        return 1;
    }
    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT cr; GetClientRect(hwnd, &cr);
        HDC mem = CreateCompatibleDC(hdc);
        HBITMAP bmp = CreateCompatibleBitmap(hdc, cr.right, cr.bottom);
        HBITMAP oldb = (HBITMAP)SelectObject(mem, bmp);
        FillRectColor(mem, cr, COL_HEADER);
        HFONT oldF = (HFONT)SelectObject(mem, g.hFontBold);
        SetBkMode(mem, TRANSPARENT);
        SetTextColor(mem, COL_TEXT);
        int n = Header_GetItemCount(hwnd);
        int x = 0;
        for (int i = 0; i < n; i++) {
            RECT ir; Header_GetItemRect(hwnd, i, &ir);
            HDITEMW hd{};
            wchar_t buf[128]; buf[0] = 0;
            hd.mask = HDI_TEXT | HDI_FORMAT;
            hd.pszText = buf;
            hd.cchTextMax = 128;
            Header_GetItem(hwnd, i, &hd);
            DrawVLine(mem, ir.right - 1, ir.top + 4, ir.bottom - 4, COL_BORDER);
            UINT flags = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
            flags |= ((hd.fmt & HDF_RIGHT) ? DT_RIGHT : DT_LEFT);
            RECT tr = { ir.left + 8, ir.top, ir.right - 8, ir.bottom };
            DrawTextW(mem, buf, -1, &tr, flags);
            if (hd.fmt & HDF_SORTUP) {
                int cx = ir.right - 14, cy = (ir.top + ir.bottom) / 2 - 2;
                POINT pts[3] = { {cx, cy + 4}, {cx + 4, cy + 4}, {cx + 2, cy} };
                HBRUSH bb = CreateSolidBrush(COL_DIM);
                HBRUSH ob = (HBRUSH)SelectObject(mem, bb);
                Polygon(mem, pts, 3);
                SelectObject(mem, ob);
                DeleteObject(bb);
            } else if (hd.fmt & HDF_SORTDOWN) {
                int cx = ir.right - 14, cy = (ir.top + ir.bottom) / 2 - 2;
                POINT pts[3] = { {cx, cy}, {cx + 4, cy}, {cx + 2, cy + 4} };
                HBRUSH bb = CreateSolidBrush(COL_DIM);
                HBRUSH ob = (HBRUSH)SelectObject(mem, bb);
                Polygon(mem, pts, 3);
                SelectObject(mem, ob);
                DeleteObject(bb);
            }
            x = ir.right;
        }
        DrawHLine(mem, 0, cr.right, cr.bottom - 1, COL_BORDER);
        SelectObject(mem, oldF);
        BitBlt(hdc, 0, 0, cr.right, cr.bottom, mem, 0, 0, SRCCOPY);
        SelectObject(mem, oldb);
        DeleteObject(bmp);
        DeleteDC(mem);
        EndPaint(hwnd, &ps);
        return 0;
    }
    return DefSubclassProc(hwnd, msg, w, l);
}

static LRESULT CALLBACK ListSubclass(HWND hwnd, UINT msg, WPARAM w, LPARAM l, UINT_PTR id, DWORD_PTR ref) {
    Pane& p = *(Pane*)ref;
    switch (msg) {
        case WM_GETDLGCODE:
            return DLGC_WANTARROWS | DLGC_WANTCHARS;
        case WM_KEYDOWN: {
            bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            if (ctrl && w == 'A') { DoSelectAll(p); return 0; }
            if (ctrl && w == 'C') { DoCopy(p, false); return 0; }
            if (ctrl && w == 'X') { DoCopy(p, true); return 0; }
            if (ctrl && w == 'V') { DoPaste(p); return 0; }
            if (ctrl && w == 'H') { g.showHidden = !g.showHidden; for (auto& pa : g.panes) RefreshPane(pa); return 0; }
            if (ctrl && w == 'T') { if (!p.tabs.empty()) NewTab(p, p.tabs[p.activeTab].path); return 0; }
            if (ctrl && w == 'W') { CloseTab(p, p.activeTab); return 0; }
            if (ctrl && w == 'L') { ShowEditPath(p); return 0; }
            if (ctrl && w == 'N') { DoNewFolder(p); return 0; }
            if (ctrl && w == 'D') { /* dark mode always on */ return 0; }
            if (ctrl && w == 'F') {
                if (p.hwndSearch) { SetFocus(p.hwndSearch); SendMessageW(p.hwndSearch, EM_SETSEL, 0, -1); }
                return 0;
            }
            if (ctrl && shift && w == 'S') { StartFolderSizeForSelection(p); return 0; }
            if (!ctrl && w == VK_LEFT && (GetKeyState(VK_MENU) & 0x8000)) { GoBack(p); return 0; }
            if (!ctrl && w == VK_RIGHT && (GetKeyState(VK_MENU) & 0x8000)) { GoForward(p); return 0; }
            if (!ctrl && w == VK_UP && (GetKeyState(VK_MENU) & 0x8000)) { GoUp(p); return 0; }
            if (!ctrl && w == VK_HOME && (GetKeyState(VK_MENU) & 0x8000)) { NavigateHome(p); return 0; }
            if (ctrl && w == VK_TAB) {
                if (!p.tabs.empty()) {
                    int n = (int)p.tabs.size();
                    p.activeTab = shift ? (p.activeTab + n - 1) % n : (p.activeTab + 1) % n;
                    RefreshPane(p);
                    InvalidateRect(p.hwnd, nullptr, FALSE);
                }
                return 0;
            }
            if (ctrl && w >= '1' && w <= '4') { ApplySort(p, (int)w - '1'); return 0; }
            // Bookmarks: Ctrl+Shift+1..9 set, Ctrl+5..9 jump
            if (ctrl && shift && w >= '1' && w <= '9') {
                int idx = (int)(w - '1');
                if (!p.tabs.empty()) {
                    g.bookmarks[idx] = p.tabs[p.activeTab].path;
                    SaveBookmarks();
                    SetWindowTextW(g.hwndStatus, (std::wstring(L"Bookmark ") + (wchar_t)w + L" saved").c_str());
                }
                return 0;
            }
            if (ctrl && !shift && w >= '5' && w <= '9') {
                int bi = (int)(w - '5');
                if (!g.bookmarks[bi].empty() && DirExists(g.bookmarks[bi]))
                    NavigateTab(p, p.activeTab, g.bookmarks[bi]);
                return 0;
            }
            // New features
            if (ctrl && w == 'Z') { DoUndo(); return 0; }
            if (ctrl && w == 'P') { ShowCmdPalette(p); return 0; }
            if (ctrl && w == 'K') { ShowQuickOpen(p); return 0; }
            if (ctrl && w == 'G') { ToggleGridView(p); return 0; }
            if (ctrl && shift && w == 'V') { ShowClipHistoryDialog(p); return 0; }
            if (ctrl && shift && w == 'H') { ToggleHexView(p); return 0; }
            if (ctrl && shift && w == 'G') { if (!p.tabs.empty()) StartGitStatus(p.tabs[p.activeTab].path); return 0; }
            if (w == VK_TAB) {
                if (g.showPane2) {
                    int other = 1 - p.paneIndex;
                    SetActivePane(other);
                    SetFocus(g.panes[other].hwndList);
                }
                return 0;
            }
            if (w == VK_F5) { RefreshPane(p); return 0; }
            if (w == VK_F1) { SetFocus(g.hwndSidebar); return 0; }
            if (w == VK_F3) {
                g.showPane2 = !g.showPane2;
                if (g.showPane2) {
                    ShowWindow(g.panes[1].hwnd, SW_SHOW);
                    RefreshPane(g.panes[1]);
                } else {
                    g.activePane = 0;
                    SetFocus(g.panes[0].hwndList);
                }
                LayoutMain(g.hwndMain);
                return 0;
            }
            break;
        }
        case WM_CHAR: {
            // Keyboard jump-to-item
            wchar_t ch = (wchar_t)w;
            if (ch >= 32 && ch < 127) {
                static wchar_t jumpBuf[32] = {};
                static DWORD jumpTick = 0;
                DWORD now = GetTickCount();
                if (now - jumpTick > 1000) jumpBuf[0] = 0;
                jumpTick = now;
                int jlen = (int)wcslen(jumpBuf);
                if (jlen < 30) { jumpBuf[jlen] = ch; jumpBuf[jlen+1] = 0; }
                int count = EffectiveItemCount(p);
                for (int i = 0; i < count; i++) {
                    FileEntry* e = EffectiveItem(p, i);
                    if (e && _wcsnicmp(e->name.c_str(), jumpBuf, wcslen(jumpBuf)) == 0) {
                        ListView_SetItemState(p.hwndList, -1, 0, LVIS_SELECTED|LVIS_FOCUSED);
                        ListView_SetItemState(p.hwndList, i, LVIS_SELECTED|LVIS_FOCUSED, LVIS_SELECTED|LVIS_FOCUSED);
                        ListView_EnsureVisible(p.hwndList, i, FALSE);
                        break;
                    }
                }
                return 0;
            }
            break;
        }
        case WM_LBUTTONDOWN: {
            SetActivePane(p.paneIndex);
            break;
        }
        case WM_SETFOCUS: {
            SetActivePane(p.paneIndex);
            break;
        }
        case WM_MOUSEWHEEL: {
            int delta = GET_WHEEL_DELTA_WPARAM(w);
            p.scrollVelocity -= (float)delta * 0.5f;
            SetTimer(hwnd, TIMER_SCROLL + p.paneIndex, 8, nullptr);
            return 0; // suppress default jump scroll
        }
        case WM_TIMER: {
            if ((int)w == TIMER_SCROLL + p.paneIndex) {
                if (std::fabs(p.scrollVelocity) < 1.5f) {
                    p.scrollVelocity = 0.f;
                    KillTimer(hwnd, w);
                } else {
                    int step = (int)(p.scrollVelocity * 0.3f);
                    if (step == 0) step = (p.scrollVelocity > 0.f) ? 1 : -1;
                    ListView_Scroll(hwnd, 0, step);
                    p.scrollVelocity *= 0.78f;
                }
                return 0;
            }
            break;
        }
    }
    return DefSubclassProc(hwnd, msg, w, l);
}

static LRESULT CALLBACK PathEditSubclass(HWND hwnd, UINT msg, WPARAM w, LPARAM l, UINT_PTR id, DWORD_PTR ref) {
    Pane& p = *(Pane*)ref;
    if (msg == WM_KEYDOWN) {
        if (w == VK_RETURN) {
            wchar_t buf[1024];
            GetWindowTextW(hwnd, buf, 1024);
            ShowWindow(hwnd, SW_HIDE);
            SetFocus(p.hwndList);
            std::wstring np = NormalizePath(buf);
            if (DirExists(np)) NavigateTab(p, p.activeTab, np);
            else MessageBeep(MB_ICONHAND);
            return 0;
        }
        if (w == VK_ESCAPE) {
            ShowWindow(hwnd, SW_HIDE);
            SetFocus(p.hwndList);
            return 0;
        }
    }
    if (msg == WM_CHAR && (w == VK_RETURN || w == VK_ESCAPE)) return 0;
    return DefSubclassProc(hwnd, msg, w, l);
}

static LRESULT CALLBACK RenameEditSubclass(HWND hwnd, UINT msg, WPARAM w, LPARAM l, UINT_PTR idx, DWORD_PTR ref) {
    Pane& p = *(Pane*)ref;
    auto finish = [&](bool commit) {
        if (!IsWindow(hwnd)) return;
        wchar_t buf[MAX_PATH];
        GetWindowTextW(hwnd, buf, MAX_PATH);
        std::wstring newName = buf;
        RemoveWindowSubclass(hwnd, RenameEditSubclass, idx);
        DestroyWindow(hwnd);
        SetFocus(p.hwndList);
        if (commit && !newName.empty() && !p.tabs.empty()) {
            FileEntry* ep = EffectiveItem(p, (int)idx);
            if (ep && newName != ep->name) {
                std::wstring oldP = CombinePath(p.tabs[p.activeTab].path, ep->name);
                if (RenameFile(g.hwndMain, oldP, newName)) RefreshPane(p);
            }
        }
    };
    if (msg == WM_KEYDOWN) {
        if (w == VK_RETURN) { finish(true); return 0; }
        if (w == VK_ESCAPE) { finish(false); return 0; }
    }
    if (msg == WM_CHAR && (w == VK_RETURN || w == VK_ESCAPE)) return 0;
    if (msg == WM_KILLFOCUS) {
        finish(true);
        return 0;
    }
    return DefSubclassProc(hwnd, msg, w, l);
}

static LRESULT CALLBACK SearchEditSubclass(HWND hwnd, UINT msg, WPARAM w, LPARAM l, UINT_PTR id, DWORD_PTR ref) {
    Pane& p = *(Pane*)ref;
    if (msg == WM_KEYDOWN) {
        if (w == VK_ESCAPE) {
            SetWindowTextW(hwnd, L"");
            if (!p.tabs.empty()) {
                p.tabs[p.activeTab].search.clear();
                p.tabs[p.activeTab].filteredEntries.clear();
                ListView_SetItemCountEx(p.hwndList, EffectiveItemCount(p),
                    LVSICF_NOSCROLL | LVSICF_NOINVALIDATEALL);
                InvalidateRect(p.hwndList, nullptr, TRUE);
            }
            SetFocus(p.hwndList);
            return 0;
        }
        if (w == VK_RETURN || w == VK_DOWN) {
            SetFocus(p.hwndList);
            return 0;
        }
    }
    if (msg == WM_CHAR && (w == VK_RETURN || w == VK_ESCAPE)) return 0;
    return DefSubclassProc(hwnd, msg, w, l);
}

// =====================================================================
// CONTEXT MENU
// =====================================================================

static void ShowContextMenu(Pane& p, POINT screenPt) {
    auto sel = GetSelectedPaths(p);
    HMENU m = CreatePopupMenu();
    bool hasSel = !sel.empty();
    bool isFolder = false;
    if (hasSel && sel.size() == 1) {
        DWORD a = GetFileAttributesW(sel[0].c_str());
        isFolder = (a != INVALID_FILE_ATTRIBUTES) && (a & FILE_ATTRIBUTE_DIRECTORY);
    }
    bool clipHas = !GetClipboardFiles().empty();

    if (hasSel) {
        AppendMenuW(m, MF_STRING, IDM_OPEN, L"Open");
        if (!isFolder)
            AppendMenuW(m, MF_STRING, IDM_OPEN_WITH, L"Open With...");
        if (isFolder) {
            AppendMenuW(m, MF_STRING, IDM_OPEN_NEW_TAB, L"Open in New Tab");
            AppendMenuW(m, MF_STRING, IDM_OPEN_OTHER_PANE, L"Open in Other Pane");
            AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
            AppendMenuW(m, MF_STRING, IDM_PIN_SIDEBAR,
                IsPinned(sel[0]) ? L"Unpin from Sidebar" : L"Pin to Sidebar");
        }
        AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(m, MF_STRING, IDM_COPY, L"Copy\tCtrl+C");
        AppendMenuW(m, MF_STRING, IDM_CUT, L"Cut\tCtrl+X");
    }
    AppendMenuW(m, MF_STRING | (clipHas ? 0 : MF_GRAYED), IDM_PASTE, L"Paste\tCtrl+V");
    if (hasSel) {
        AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(m, MF_STRING, IDM_DELETE, L"Delete\tDel");
        AppendMenuW(m, MF_STRING, IDM_RENAME, L"Rename\tF2");
        if (sel.size() > 1)
            AppendMenuW(m, MF_STRING, IDM_BULK_RENAME, L"Bulk Rename...");
        AppendMenuW(m, MF_STRING, IDM_FOLDER_SIZE, L"Folder Size\tCtrl+Shift+S");
    }
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    // Copy Path submenu
    HMENU mCopyPath = CreatePopupMenu();
    AppendMenuW(mCopyPath, MF_STRING, IDM_COPY_PATH_FULL, L"Full Path");
    AppendMenuW(mCopyPath, MF_STRING, IDM_COPY_PATH_NAME, L"Filename");
    AppendMenuW(mCopyPath, MF_STRING, IDM_COPY_PATH_DIR,  L"Directory");
    AppendMenuW(m, MF_POPUP, (UINT_PTR)mCopyPath, L"Copy Path");
    // Selection
    AppendMenuW(m, MF_STRING, IDM_SELECT_PATTERN, L"Select by Pattern...");
    AppendMenuW(m, MF_STRING, IDM_INVERT_SEL,     L"Invert Selection");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    // File ops
    AppendMenuW(m, MF_STRING, IDM_ZIP_SELECTED, L"Compress to .zip");
    if (hasSel && sel.size() == 1) {
        std::wstring selExt = sel[0];
        // uppercase ext check
        size_t dp = selExt.rfind(L'.');
        std::wstring ext = (dp != std::wstring::npos) ? selExt.substr(dp+1) : L"";
        for (auto& c : ext) c = towupper(c);
        if (ext == L"ZIP") {
            AppendMenuW(m, MF_STRING, IDM_UNZIP_HERE,        L"Extract Here");
            AppendMenuW(m, MF_STRING, IDM_UNZIP_TO_FOLDER,  L"Extract to Folder...");
        }
        if (ext == L"RAR") {
            AppendMenuW(m, MF_STRING, IDM_UNRAR_HERE,        L"Extract RAR Here");
            AppendMenuW(m, MF_STRING, IDM_UNRAR_TO_FOLDER,  L"Extract RAR to Folder...");
        }
    }
    if (hasSel && !isFolder) {
        AppendMenuW(m, MF_STRING, IDM_SPLIT_FILE, L"Split File...");
        AppendMenuW(m, MF_STRING, IDM_JOIN_FILES,  L"Join Files");
    }
    AppendMenuW(m, MF_STRING, IDM_FILE_DIFF,      L"File Diff");
    AppendMenuW(m, MF_STRING, IDM_FOLDER_DIFF,    L"Folder Diff");
    AppendMenuW(m, MF_STRING, IDM_FIND_DUPES,     L"Find Duplicates...");
    if (hasSel && !isFolder)
        AppendMenuW(m, MF_STRING, IDM_PROC_LOCK,  L"Check File Lock");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    // Color tags
    AppendMenuW(m, MF_STRING, IDM_COPY_PATH,      L"Tag Color...");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING | (g.showPane2 ? MF_CHECKED : 0), IDM_COMPUTE_HASH, L"Dual Pane\tF3");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING, IDM_NEW_FOLDER,     L"New Folder\tCtrl+N");
    AppendMenuW(m, MF_STRING, IDM_REFRESH,        L"Refresh\tF5");
    AppendMenuW(m, MF_STRING, IDM_OPEN_TERMINAL,  L"Open Terminal Here");
    AppendMenuW(m, MF_STRING, IDM_CONNECT_SERVER, L"Connect to Network Share...");
    AppendMenuW(m, MF_STRING | (g.showHidden ? MF_CHECKED : 0), IDM_TOGGLE_HIDDEN, L"Show Hidden Files\tCtrl+H");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING | (g.viewMode[p.paneIndex] ? MF_CHECKED : 0), IDM_GRID_VIEW, L"Grid View\tCtrl+G");
    AppendMenuW(m, MF_STRING, IDM_GIT_STATUS,     L"Refresh Git Status");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING, IDM_UNDO,           L"Undo\tCtrl+Z");
    AppendMenuW(m, MF_STRING, IDM_SETTINGS,       L"Settings...");
    AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(m, MF_STRING, IDM_PROPERTIES,     L"Properties");
    AppendMenuW(m, MF_STRING, IDM_SHELL_MENU, L"More options (Shell)\tShift+RClick");

    int cmd = TrackPopupMenu(m, TPM_RETURNCMD | TPM_RIGHTBUTTON, screenPt.x, screenPt.y, 0, p.hwnd, nullptr);
    DestroyMenu(m);
    switch (cmd) {
        case IDM_OPEN: OpenSelected(p); break;
        case IDM_OPEN_NEW_TAB: if (hasSel && isFolder) NewTab(p, sel[0]); break;
        case IDM_OPEN_OTHER_PANE: if (hasSel && isFolder) {
            int o = 1 - p.paneIndex;
            NavigateTab(g.panes[o], g.panes[o].activeTab, sel[0]);
        } break;
        case IDM_PIN_SIDEBAR: if (hasSel && isFolder) TogglePin(sel[0]); break;
        case IDM_COPY: DoCopy(p, false); break;
        case IDM_CUT: DoCopy(p, true); break;
        case IDM_PASTE: DoPaste(p); break;
        case IDM_DELETE: DoDelete(p, false); break;
        case IDM_RENAME: DoRename(p); break;
        case IDM_BULK_RENAME: ShowBulkRenameDialog(p); break;
        case IDM_FOLDER_SIZE: StartFolderSizeForSelection(p); break;
        case IDM_COPY_PATH_FULL: DoCopyPath(p, 0); break;
        case IDM_COPY_PATH_NAME: DoCopyPath(p, 1); break;
        case IDM_COPY_PATH_DIR:  DoCopyPath(p, 2); break;
        case IDM_SELECT_PATTERN: DoSelectByPattern(p); break;
        case IDM_INVERT_SEL:     DoInvertSelection(p); break;
        case IDM_ZIP_SELECTED:      DoZipSelected(p); break;
        case IDM_UNZIP_HERE:        if (hasSel) DoExtractZip(p, sel[0], false); break;
        case IDM_UNZIP_TO_FOLDER:   if (hasSel) DoExtractZip(p, sel[0], true); break;
        case IDM_UNRAR_HERE:        if (hasSel) DoExtractRar(p, sel[0], false); break;
        case IDM_UNRAR_TO_FOLDER:   if (hasSel) DoExtractRar(p, sel[0], true); break;
        case IDM_SPLIT_FILE:     DoSplitFile(p); break;
        case IDM_JOIN_FILES:     DoJoinFiles(p); break;
        case IDM_FILE_DIFF:      DoFileDiff(p); break;
        case IDM_FOLDER_DIFF:    DoFolderDiff(p); break;
        case IDM_FIND_DUPES:     DoFindDuplicates(p); break;
        case IDM_PROC_LOCK:      if (hasSel) ShowFileLockInfo(sel[0]); break;
        case IDM_COPY_PATH:      SetColorTag(p); break;
        case IDM_OPEN_TERMINAL:  DoOpenTerminal(p); break;
        case IDM_CONNECT_SERVER: DoConnectServer(p); break;
        case IDM_GRID_VIEW:      ToggleGridView(p); break;
        case IDM_GIT_STATUS:     if (!p.tabs.empty()) StartGitStatus(p.tabs[p.activeTab].path); break;
        case IDM_UNDO:           DoUndo(); break;
        case IDM_SETTINGS:       ShowSettingsDialog(); break;
        case IDM_COMPUTE_HASH:
            g.showPane2 = !g.showPane2;
            if (g.showPane2) { ShowWindow(g.panes[1].hwnd, SW_SHOW); RefreshPane(g.panes[1]); }
            else { g.activePane = 0; SetFocus(g.panes[0].hwndList); }
            LayoutMain(g.hwndMain);
            break;
        case IDM_NEW_FOLDER: DoNewFolder(p); break;
        case IDM_REFRESH: RefreshPane(p); break;
        case IDM_SORT_NAME: ApplySort(p, 0); break;
        case IDM_SORT_SIZE: ApplySort(p, 1); break;
        case IDM_SORT_DATE: ApplySort(p, 2); break;
        case IDM_SORT_TYPE: ApplySort(p, 3); break;
        case IDM_TOGGLE_HIDDEN:
            g.showHidden = !g.showHidden;
            for (auto& pa : g.panes) RefreshPane(pa);
            break;
        case IDM_PROPERTIES:
            if (hasSel) ShowProperties(g.hwndMain, sel[0]);
            break;
        case IDM_OPEN_WITH:
            if (hasSel) ShellExecuteW(g.hwndMain, L"openwith", sel[0].c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            break;
        case IDM_SHELL_MENU:
            ShowShellContextMenuForPane(p, screenPt);
            break;
    }
}

// =====================================================================
// BULK RENAME DIALOG
// =====================================================================

struct BulkRenameCtx {
    Pane* pane;
    std::vector<std::wstring> names;
    std::vector<std::wstring> preview;
    HWND dlg;
};
static BulkRenameCtx s_brc;

static std::wstring ApplyRenamePattern(const std::wstring& orig, const std::wstring& find,
                                       const std::wstring& repl, bool numbering, int n) {
    std::wstring result = orig;
    if (!find.empty()) {
        size_t pos = 0;
        while ((pos = result.find(find, pos)) != std::wstring::npos) {
            result.replace(pos, find.size(), repl);
            pos += repl.size();
        }
    }
    if (numbering) {
        size_t dot = result.find_last_of(L'.');
        std::wstring num = std::to_wstring(n);
        if (dot != std::wstring::npos && dot > 0)
            result = result.substr(0, dot) + L"_" + num + result.substr(dot);
        else
            result = result + L"_" + num;
    }
    return result;
}

static void BulkRenameRefreshPreview(HWND dlg) {
    wchar_t find[512]={}, repl[512]={};
    GetDlgItemTextW(dlg, ID_BULK_FIND, find, 511);
    GetDlgItemTextW(dlg, ID_BULK_REPL, repl, 511);
    bool numbering = IsDlgButtonChecked(dlg, ID_BULK_NUMBER) == BST_CHECKED;
    HWND lb = GetDlgItem(dlg, ID_BULK_LIST);
    SendMessageW(lb, LB_RESETCONTENT, 0, 0);
    s_brc.preview.clear();
    for (int i = 0; i < (int)s_brc.names.size(); i++) {
        std::wstring pv = ApplyRenamePattern(s_brc.names[i], find, repl, numbering, i+1);
        s_brc.preview.push_back(pv);
        std::wstring line = s_brc.names[i] + L"  →  " + pv;
        SendMessageW(lb, LB_ADDSTRING, 0, (LPARAM)line.c_str());
    }
}

static INT_PTR CALLBACK BulkRenameDlgProc(HWND dlg, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_INITDIALOG: {
            s_brc.dlg = dlg;
            BulkRenameRefreshPreview(dlg);
            return TRUE;
        }
        case WM_COMMAND: {
            int id = LOWORD(w), code = HIWORD(w);
            if ((id == ID_BULK_FIND || id == ID_BULK_REPL) && code == EN_CHANGE)
                BulkRenameRefreshPreview(dlg);
            if (id == ID_BULK_NUMBER && (code == BN_CLICKED || code == 0))
                BulkRenameRefreshPreview(dlg);
            if (id == ID_BULK_OK) {
                for (int i = 0; i < (int)s_brc.names.size(); i++) {
                    if (s_brc.preview[i].empty() || s_brc.preview[i] == s_brc.names[i]) continue;
                    std::wstring oldP = CombinePath(s_brc.pane->tabs[s_brc.pane->activeTab].path, s_brc.names[i]);
                    RenameFile(GetParent(dlg), oldP, s_brc.preview[i]);
                }
                RefreshPane(*s_brc.pane);
                EndDialog(dlg, 1);
            }
            if (id == ID_BULK_CANCEL) EndDialog(dlg, 0);
            return TRUE;
        }
        case WM_CLOSE:
            EndDialog(dlg, 0);
            return TRUE;
    }
    return FALSE;
}

static LRESULT CALLBACK BulkRenameWndProc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
    if (BulkRenameDlgProc(h, m, wp, lp)) return 0;
    return DefWindowProcW(h, m, wp, lp);
}

static void ShowBulkRenameDialog(Pane& p) {
    auto sel = GetSelectedPaths(p);
    if (sel.size() < 2) { MessageBeep(MB_ICONINFORMATION); return; }
    s_brc.pane = &p;
    s_brc.names.clear();
    for (auto& sp : sel) s_brc.names.push_back(NameFromPath(sp));

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"STATIC", L"",
        WS_POPUP | WS_CAPTION | WS_SYSMENU, 0, 0, 540, 440, g.hwndMain, nullptr, g.hInst, nullptr);
    SetWindowTextW(dlg, L"Bulk Rename");

    auto addLabel = [&](const wchar_t* txt, int x, int y, int w2, int h2) {
        HWND lbl = CreateWindowExW(0, L"STATIC", txt, WS_CHILD|WS_VISIBLE, x,y,w2,h2, dlg, nullptr, g.hInst, nullptr);
        SendMessageW(lbl, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    };
    auto addEdit = [&](int id, int x, int y, int w2, int h2) -> HWND {
        HWND ed = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL, x,y,w2,h2, dlg, (HMENU)(INT_PTR)id, g.hInst, nullptr);
        SendMessageW(ed, WM_SETFONT, (WPARAM)g.hFont, TRUE);
        return ed;
    };
    auto addBtn = [&](int id, const wchar_t* txt, int x, int y, int w2, int h2) {
        HWND btn = CreateWindowExW(0, L"BUTTON", txt, WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, x,y,w2,h2, dlg, (HMENU)(INT_PTR)id, g.hInst, nullptr);
        SendMessageW(btn, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    };

    addLabel(L"Find:", 12, 14, 60, 20);
    addEdit(ID_BULK_FIND, 75, 12, 220, 22);
    addLabel(L"Replace:", 310, 14, 60, 20);
    addEdit(ID_BULK_REPL, 375, 12, 150, 22);

    HWND chk = CreateWindowExW(0, L"BUTTON", L"Add number", WS_CHILD|WS_VISIBLE|BS_AUTOCHECKBOX,
        12, 44, 150, 22, dlg, (HMENU)(INT_PTR)ID_BULK_NUMBER, g.hInst, nullptr);
    SendMessageW(chk, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    HWND lb = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_NOINTEGRALHEIGHT|LBS_HASSTRINGS,
        12, 74, 510, 300, dlg, (HMENU)(INT_PTR)ID_BULK_LIST, g.hInst, nullptr);
    SendMessageW(lb, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    addBtn(ID_BULK_OK,     L"Apply", 334, 386, 80, 28);
    addBtn(ID_BULK_CANCEL, L"Cancel", 422, 386, 80, 28);

    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)BulkRenameWndProc);

    // Init preview
    s_brc.dlg = dlg;
    BulkRenameRefreshPreview(dlg);

    RECT rc; GetWindowRect(g.hwndMain, &rc);
    int dx = rc.left + (rc.right - rc.left - 540) / 2;
    int dy = rc.top + (rc.bottom - rc.top - 440) / 2;
    SetWindowPos(dlg, nullptr, dx, dy, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
    ShowWindow(dlg, SW_SHOW);

    MSG msg;
    while (IsWindow(dlg) && GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(dlg, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

static void ShowShellContextMenuForPane(Pane& p, POINT screenPt) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) {
        if (p.tabs.empty()) return;
        sel.push_back(p.tabs[p.activeTab].path);
    }
    if (ShowShellContextMenuFor(p.hwnd, sel, screenPt))
        RefreshPane(p);
}

// =====================================================================
// SIMPLE INPUT BOX
// =====================================================================

struct InputBoxCtx { std::wstring result; const wchar_t* prompt; };
static InputBoxCtx s_ibCtx;
static HWND s_ibEdit = nullptr;

static INT_PTR CALLBACK InputBoxDlgProc(HWND dlg, UINT msg, WPARAM w, LPARAM /*l*/) {
    switch (msg) {
        case WM_INITDIALOG: {
            SetDlgItemTextW(dlg, 100, s_ibCtx.prompt);
            s_ibEdit = GetDlgItem(dlg, 101);
            SetWindowTextW(s_ibEdit, s_ibCtx.result.c_str());
            SendMessageW(s_ibEdit, EM_SETSEL, 0, -1);
            SetFocus(s_ibEdit);
            return FALSE;
        }
        case WM_COMMAND:
            if (LOWORD(w) == IDOK) {
                wchar_t buf[1024] = {};
                GetWindowTextW(s_ibEdit, buf, 1023);
                s_ibCtx.result = buf;
                EndDialog(dlg, 1);
            } else if (LOWORD(w) == IDCANCEL) {
                s_ibCtx.result.clear();
                EndDialog(dlg, 0);
            }
            return TRUE;
        case WM_CLOSE:
            s_ibCtx.result.clear();
            EndDialog(dlg, 0);
            return TRUE;
    }
    return FALSE;
}

static LRESULT CALLBACK InputBoxWndProc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
    if (InputBoxDlgProc(h, m, wp, lp)) return 0;
    return DefWindowProcW(h, m, wp, lp);
}

static std::wstring SimpleInputBox(HWND parent, const wchar_t* title, const wchar_t* prompt, const wchar_t* defVal) {
    s_ibCtx.prompt = prompt;
    s_ibCtx.result = defVal ? defVal : L"";

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"STATIC", title,
        WS_POPUP | WS_CAPTION | WS_SYSMENU,
        0, 0, 400, 140, parent, nullptr, g.hInst, nullptr);
    if (!dlg) return L"";

    SetWindowTextW(dlg, title);

    HWND lbl = CreateWindowExW(0, L"STATIC", prompt, WS_CHILD|WS_VISIBLE, 12, 14, 370, 20, dlg, (HMENU)100, g.hInst, nullptr);
    SendMessageW(lbl, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    s_ibEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", defVal ? defVal : L"",
        WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL, 12, 38, 370, 24, dlg, (HMENU)101, g.hInst, nullptr);
    SendMessageW(s_ibEdit, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    SendMessageW(s_ibEdit, EM_SETSEL, 0, -1);

    HWND ok = CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD|WS_VISIBLE|BS_DEFPUSHBUTTON,
        220, 76, 74, 26, dlg, (HMENU)IDOK, g.hInst, nullptr);
    SendMessageW(ok, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    HWND cn = CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
        304, 76, 74, 26, dlg, (HMENU)IDCANCEL, g.hInst, nullptr);
    SendMessageW(cn, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)InputBoxWndProc);

    RECT rc; GetWindowRect(parent, &rc);
    int dx = rc.left + (rc.right - rc.left - 400) / 2;
    int dy = rc.top + (rc.bottom - rc.top - 140) / 2;
    SetWindowPos(dlg, nullptr, dx, dy, 0, 0, SWP_NOSIZE|SWP_NOZORDER);
    ShowWindow(dlg, SW_SHOW);
    SetFocus(s_ibEdit);

    MSG msg;
    while (IsWindow(dlg) && GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(dlg, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    return s_ibCtx.result;
}

// =====================================================================
// SETTINGS SAVE / LOAD
// =====================================================================

static void SaveSettings() {
    auto f = SettingsFile();
    if (f.empty()) return;
    HANDLE h = CreateFileW(f.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;

    auto writeLine = [&](const std::wstring& line) {
        std::string s;
        int n = WideCharToMultiByte(CP_UTF8, 0, line.c_str(), (int)line.size(), nullptr, 0, nullptr, nullptr);
        if (n > 0) {
            s.resize(n);
            WideCharToMultiByte(CP_UTF8, 0, line.c_str(), (int)line.size(), s.data(), n, nullptr, nullptr);
        }
        s += "\n";
        DWORD wrote;
        WriteFile(h, s.c_str(), (DWORD)s.size(), &wrote, nullptr);
    };

    writeLine(std::wstring(L"font_size=") + std::to_wstring(g.settingFontSize));
    writeLine(std::wstring(L"sidebar_w=") + std::to_wstring(g.sidebarW));
    writeLine(std::wstring(L"preview_w=") + std::to_wstring(g.previewW));
    writeLine(std::wstring(L"show_hidden=") + (g.showHidden ? L"1" : L"0"));
    writeLine(std::wstring(L"auto_refresh=") + (g.settingAutoRefresh ? L"1" : L"0"));
    writeLine(std::wstring(L"start_path=") + g.settingStartPath);
    for (int i = 0; i < 9; i++)
        writeLine(std::wstring(L"bookmark_") + std::to_wstring(i+1) + L"=" + g.bookmarks[i]);
    for (auto& kv : g.customAssoc)
        writeLine(std::wstring(L"assoc_") + kv.first + L"=" + kv.second);
    for (auto& kv : g.colorTags) {
        wchar_t buf[32];
        StringCchPrintfW(buf, 32, L"%06X", (unsigned)kv.second);
        writeLine(std::wstring(L"colortag_") + kv.first + L"=" + buf);
    }
    CloseHandle(h);
}

static void LoadSettings() {
    auto f = SettingsFile();
    if (f.empty()) return;
    HANDLE h = CreateFileW(f.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    LARGE_INTEGER sz{};
    GetFileSizeEx(h, &sz);
    if (sz.QuadPart <= 0 || sz.QuadPart > 512*1024) { CloseHandle(h); return; }
    std::string raw((size_t)sz.QuadPart, '\0');
    DWORD got = 0;
    ReadFile(h, raw.data(), (DWORD)sz.QuadPart, &got, nullptr);
    CloseHandle(h);

    int wn = MultiByteToWideChar(CP_UTF8, 0, raw.data(), got, nullptr, 0);
    if (wn <= 0) return;
    std::wstring all(wn, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, raw.data(), got, all.data(), wn);

    size_t s = 0;
    for (size_t i = 0; i <= all.size(); i++) {
        if (i == all.size() || all[i] == L'\n' || all[i] == L'\r') {
            if (i > s) {
                std::wstring line = all.substr(s, i - s);
                while (!line.empty() && (line.back() == L'\r' || line.back() == L'\n')) line.pop_back();
                size_t eq = line.find(L'=');
                if (eq != std::wstring::npos) {
                    std::wstring key = line.substr(0, eq);
                    std::wstring val = line.substr(eq + 1);
                    if (key == L"font_size") { int v = _wtoi(val.c_str()); if (v >= 8 && v <= 24) g.settingFontSize = v; }
                    else if (key == L"sidebar_w") { int v = _wtoi(val.c_str()); if (v >= 60 && v <= 600) g.sidebarW = v; }
                    else if (key == L"preview_w") { int v = _wtoi(val.c_str()); if (v >= 60 && v <= 800) g.previewW = v; }
                    else if (key == L"show_hidden") g.showHidden = (val == L"1");
                    else if (key == L"auto_refresh") g.settingAutoRefresh = (val == L"1");
                    else if (key == L"start_path") g.settingStartPath = val;
                    else if (key.size() > 9 && key.substr(0,9) == L"bookmark_") {
                        int idx = _wtoi(key.substr(9).c_str()) - 1;
                        if (idx >= 0 && idx < 9) g.bookmarks[idx] = val;
                    }
                    else if (key.size() > 6 && key.substr(0,6) == L"assoc_") {
                        std::wstring ext = key.substr(6);
                        for (auto& c : ext) c = (wchar_t)towupper(c);
                        g.customAssoc[ext] = val;
                    }
                    else if (key.size() > 9 && key.substr(0,9) == L"colortag_") {
                        std::wstring path = key.substr(9);
                        unsigned int rgb = 0;
                        swscanf_s(val.c_str(), L"%X", &rgb);
                        g.colorTags[path] = (COLORREF)rgb;
                    }
                }
            }
            s = i + 1;
        }
    }
}

static void SaveBookmarks() { SaveSettings(); }
static void LoadBookmarks() { /* loaded in LoadSettings */ }

// =====================================================================
// SESSION SAVE / LOAD
// =====================================================================

static void SaveSession() {
    auto f = SessionFile();
    if (f.empty()) return;
    HANDLE h = CreateFileW(f.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    const unsigned char bom[2] = {0xFF, 0xFE};
    DWORD wrote;
    WriteFile(h, bom, 2, &wrote, nullptr);
    for (int pi = 0; pi < 2; pi++) {
        for (auto& tab : g.panes[pi].tabs) {
            WriteFile(h, tab.path.data(), (DWORD)(tab.path.size()*sizeof(wchar_t)), &wrote, nullptr);
            wchar_t nl = L'\n';
            WriteFile(h, &nl, sizeof(wchar_t), &wrote, nullptr);
        }
        std::wstring sep = L"---\n";
        WriteFile(h, sep.data(), (DWORD)(sep.size()*sizeof(wchar_t)), &wrote, nullptr);
    }
    CloseHandle(h);
}

static void LoadSession() {
    auto f = SessionFile();
    if (f.empty()) return;
    HANDLE h = CreateFileW(f.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    LARGE_INTEGER sz{};
    GetFileSizeEx(h, &sz);
    if (sz.QuadPart <= 0 || sz.QuadPart > 512*1024) { CloseHandle(h); return; }
    std::vector<uint8_t> raw((size_t)sz.QuadPart);
    DWORD got = 0;
    ReadFile(h, raw.data(), (DWORD)sz.QuadPart, &got, nullptr);
    CloseHandle(h);
    std::wstring all;
    size_t off = 0;
    if (got >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) {
        off = 2;
        all.assign((wchar_t*)(raw.data()+off), (got-off)/2);
    } else {
        int wn = MultiByteToWideChar(CP_UTF8, 0, (char*)raw.data(), got, nullptr, 0);
        if (wn > 0) { all.resize(wn); MultiByteToWideChar(CP_UTF8, 0, (char*)raw.data(), got, all.data(), wn); }
    }

    int paneIdx = 0;
    std::vector<std::wstring> panePaths[2];
    size_t s = 0;
    for (size_t i = 0; i <= all.size(); i++) {
        if (i == all.size() || all[i] == L'\n' || all[i] == L'\r') {
            if (i > s) {
                std::wstring line = all.substr(s, i-s);
                while (!line.empty() && (line.back()==L'\r'||line.back()==L'\n')) line.pop_back();
                if (line == L"---") { if (paneIdx < 1) paneIdx++; }
                else if (paneIdx < 2 && DirExists(line)) panePaths[paneIdx].push_back(line);
            }
            s = i+1;
        }
    }
    for (int pi = 0; pi < 2; pi++) {
        if (!panePaths[pi].empty()) {
            g.panes[pi].tabs.clear();
            for (auto& path : panePaths[pi]) {
                PaneTab t; t.path = path;
                t.history.push_back(path); t.histPos = 0;
                g.panes[pi].tabs.push_back(t);
            }
            g.panes[pi].activeTab = 0;
        }
    }
}

// =====================================================================
// UNDO
// =====================================================================

static void PushUndo(int type, const std::wstring& src, const std::wstring& dst) {
    AppCtx::UndoOp op; op.type = type; op.src = src; op.dst = dst;
    g.undoStack.push_back(op);
    if (g.undoStack.size() > 50) g.undoStack.erase(g.undoStack.begin());
}

static void DoUndo() {
    if (g.undoStack.empty()) {
        MessageBoxW(g.hwndMain, L"Nothing to undo.", L"Undo", MB_OK|MB_ICONINFORMATION);
        return;
    }
    auto op = g.undoStack.back();
    g.undoStack.pop_back();
    switch (op.type) {
        case 0: // move src->dst: reverse = move dst->src
            MoveFiles(g.hwndMain, {op.dst}, ParentPath(op.src));
            break;
        case 1: // copy src->dst: reverse = delete dst
            DeleteFiles(g.hwndMain, {op.dst}, true);
            break;
        case 2: // rename src->dst: reverse = rename dst->src
            RenameFile(g.hwndMain, op.dst, NameFromPath(op.src));
            break;
        case 3: // mkdir dst: reverse = delete dst
            DeleteFiles(g.hwndMain, {op.dst}, true);
            break;
    }
    for (auto& pn : g.panes) RefreshPane(pn);
}

// =====================================================================
// SELECTION FEATURES
// =====================================================================

static void DoSelectByPattern(Pane& p) {
    std::wstring pat = SimpleInputBox(g.hwndMain, L"Select by Pattern", L"Pattern (e.g. *.log):", L"*");
    if (pat.empty()) return;
    int n = EffectiveItemCount(p);
    for (int i = 0; i < n; i++) {
        FileEntry* e = EffectiveItem(p, i);
        if (!e) continue;
        bool match = PathMatchSpecW(e->name.c_str(), pat.c_str());
        ListView_SetItemState(p.hwndList, i, match ? LVIS_SELECTED : 0, LVIS_SELECTED);
    }
}

static void DoInvertSelection(Pane& p) {
    int n = EffectiveItemCount(p);
    for (int i = 0; i < n; i++) {
        UINT cur = ListView_GetItemState(p.hwndList, i, LVIS_SELECTED);
        ListView_SetItemState(p.hwndList, i, cur ? 0 : LVIS_SELECTED, LVIS_SELECTED);
    }
}

// =====================================================================
// COPY PATH
// =====================================================================

static void DoCopyPath(Pane& p, int mode) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) return;
    std::wstring txt;
    for (auto& s2 : sel) {
        std::wstring piece;
        if (mode == 0) piece = s2;
        else if (mode == 1) piece = NameFromPath(s2);
        else piece = ParentPath(s2);
        if (!txt.empty()) txt += L"\r\n";
        txt += piece;
    }
    if (OpenClipboard(g.hwndMain)) {
        EmptyClipboard();
        HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, (txt.size()+1)*sizeof(wchar_t));
        if (hg) {
            wchar_t* dst = (wchar_t*)GlobalLock(hg);
            if (dst) { wcscpy(dst, txt.c_str()); GlobalUnlock(hg); }
            SetClipboardData(CF_UNICODETEXT, hg);
        }
        CloseClipboard();
    }
}

// =====================================================================
// OPEN TERMINAL
// =====================================================================

static void DoOpenTerminal(Pane& p) {
    if (p.tabs.empty()) return;
    std::wstring dir = p.tabs[p.activeTab].path;
    // Try Windows Terminal
    HINSTANCE r = ShellExecuteW(nullptr, L"open", L"wt.exe",
        (std::wstring(L"-d \"") + dir + L"\"").c_str(), nullptr, SW_SHOW);
    if ((INT_PTR)r <= 32) {
        // Fallback to cmd.exe
        ShellExecuteW(nullptr, L"open", L"cmd.exe", nullptr, dir.c_str(), SW_SHOW);
    }
}

// =====================================================================
// CONNECT TO SERVER
// =====================================================================

static void DoConnectServer(Pane& p) {
    std::wstring addr = SimpleInputBox(g.hwndMain, L"Connect to Network Share", L"Enter UNC path (e.g. \\\\server\\share):", L"\\\\");
    if (addr.empty()) return;
    if (DirExists(addr)) NavigateTab(p, p.activeTab, addr);
    else MessageBoxW(g.hwndMain, (std::wstring(L"Cannot access: ") + addr).c_str(), L"Connect", MB_OK|MB_ICONWARNING);
}

// =====================================================================
// NATIVE ZIP ENGINE  (no external libs, PKZIP 2.0 compatible store-only)
// =====================================================================

// Write a little-endian value into a byte buffer
static void ZipPut16(std::vector<uint8_t>& b, uint16_t v) {
    b.push_back((uint8_t)(v & 0xFF));
    b.push_back((uint8_t)(v >> 8));
}
static void ZipPut32(std::vector<uint8_t>& b, uint32_t v) {
    b.push_back((uint8_t)(v & 0xFF));
    b.push_back((uint8_t)((v >> 8) & 0xFF));
    b.push_back((uint8_t)((v >> 16) & 0xFF));
    b.push_back((uint8_t)(v >> 24));
}

static uint32_t Crc32Table[256];
static bool     Crc32Init = false;
static void BuildCrc32() {
    for (uint32_t i = 0; i < 256; i++) {
        uint32_t c = i;
        for (int j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >> 1)) : (c >> 1);
        Crc32Table[i] = c;
    }
    Crc32Init = true;
}
static uint32_t CalcCrc32(const uint8_t* data, size_t len) {
    if (!Crc32Init) BuildCrc32();
    uint32_t c = 0xFFFFFFFF;
    for (size_t i = 0; i < len; i++) c = Crc32Table[(c ^ data[i]) & 0xFF] ^ (c >> 8);
    return c ^ 0xFFFFFFFF;
}

struct ZipEntry {
    std::string nameUtf8;       // name stored in archive (relative, forward slashes)
    uint32_t    crc32 = 0;
    uint32_t    size  = 0;
    uint32_t    localOffset = 0;
    bool        isDir = false;
    uint16_t    dosDate = 0, dosTime = 0;
};

// Convert FILETIME to DOS date/time
static void FiletimeToDos(const FILETIME& ft, uint16_t& date, uint16_t& time) {
    SYSTEMTIME st{};
    FILETIME local{};
    FileTimeToLocalFileTime(&ft, &local);
    FileTimeToSystemTime(&local, &st);
    date = (uint16_t)(((st.wYear - 1980) << 9) | (st.wMonth << 5) | st.wDay);
    time = (uint16_t)((st.wHour << 11) | (st.wMinute << 5) | (st.wSecond >> 1));
}

// Narrow UTF-8 string from wide
static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string s(n - 1, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], n, nullptr, nullptr);
    return s;
}

// Write one local file header + data; returns true on success
static bool ZipWriteLocalEntry(HANDLE hOut, ZipEntry& ent, const std::wstring& fullPath) {
    std::vector<uint8_t> hdr;
    std::string& nm = ent.nameUtf8;
    hdr.reserve(30 + nm.size());
    // local file header signature
    hdr.push_back(0x50); hdr.push_back(0x4B); hdr.push_back(0x03); hdr.push_back(0x04);
    ZipPut16(hdr, 20);               // version needed: 2.0
    ZipPut16(hdr, 0x0800);           // general flags: UTF-8 name
    ZipPut16(hdr, 0);                // compression: store
    ZipPut16(hdr, ent.dosTime);
    ZipPut16(hdr, ent.dosDate);
    ZipPut32(hdr, ent.crc32);
    ZipPut32(hdr, ent.size);
    ZipPut32(hdr, ent.size);
    ZipPut16(hdr, (uint16_t)nm.size());
    ZipPut16(hdr, 0);                // extra len
    for (char c : nm) hdr.push_back((uint8_t)c);

    DWORD written;
    if (!WriteFile(hOut, hdr.data(), (DWORD)hdr.size(), &written, nullptr)) return false;

    if (!ent.isDir && ent.size > 0) {
        HANDLE hIn = CreateFileW(fullPath.c_str(), GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hIn == INVALID_HANDLE_VALUE) return false;
        const DWORD BUF = 65536;
        std::vector<uint8_t> buf(BUF);
        DWORD rd;
        while (ReadFile(hIn, buf.data(), BUF, &rd, nullptr) && rd > 0) {
            DWORD wr;
            if (!WriteFile(hOut, buf.data(), rd, &wr, nullptr) || wr != rd) {
                CloseHandle(hIn); return false;
            }
        }
        CloseHandle(hIn);
    }
    return true;
}

// Collect all files under a directory recursively, prefix relative to baseDir
static void ZipCollectFiles(const std::wstring& baseDir, const std::wstring& curDir,
                            const std::wstring& prefix,
                            std::vector<std::pair<ZipEntry, std::wstring>>& out)
{
    std::wstring pat = curDir + L"\\*";
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileExW(pat.c_str(), FindExInfoBasic, &fd,
        FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH);
    if (h == INVALID_HANDLE_VALUE) return;
    do {
        if (fd.cFileName[0] == L'.' && (fd.cFileName[1] == L'\0' ||
            (fd.cFileName[1] == L'.' && fd.cFileName[2] == L'\0'))) continue;
        std::wstring name(fd.cFileName);
        std::wstring fullPath = curDir + L"\\" + name;
        std::wstring relName  = prefix.empty() ? name : (prefix + L"/" + name);

        ZipEntry ent;
        ent.nameUtf8 = WideToUtf8(relName);
        ent.isDir    = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        FiletimeToDos(fd.ftLastWriteTime, ent.dosDate, ent.dosTime);

        if (ent.isDir) {
            std::string dirName = ent.nameUtf8 + "/";
            ent.nameUtf8 = dirName;
            ent.crc32 = 0; ent.size = 0;
            out.push_back({ent, fullPath});
            ZipCollectFiles(baseDir, fullPath, relName, out);
        } else {
            ULARGE_INTEGER sz{};
            sz.LowPart = fd.nFileSizeLow; sz.HighPart = fd.nFileSizeHigh;
            ent.size = (uint32_t)std::min(sz.QuadPart, (ULONGLONG)0xFFFFFFFF);
            // compute CRC
            HANDLE hf = CreateFileW(fullPath.c_str(), GENERIC_READ,
                FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
                OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (hf != INVALID_HANDLE_VALUE) {
                uint32_t crc = 0xFFFFFFFF;
                if (!Crc32Init) BuildCrc32();
                const DWORD BUF = 65536;
                std::vector<uint8_t> buf(BUF);
                DWORD rd;
                while (ReadFile(hf, buf.data(), BUF, &rd, nullptr) && rd > 0)
                    for (DWORD i = 0; i < rd; i++)
                        crc = Crc32Table[(crc ^ buf[i]) & 0xFF] ^ (crc >> 8);
                CloseHandle(hf);
                ent.crc32 = crc ^ 0xFFFFFFFF;
            }
            out.push_back({ent, fullPath});
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);
}

static bool ZipCreate(const std::wstring& zipPath,
                      const std::vector<std::wstring>& sources)
{
    HANDLE hOut = CreateFileW(zipPath.c_str(), GENERIC_WRITE, 0, nullptr,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hOut == INVALID_HANDLE_VALUE) return false;

    std::vector<std::pair<ZipEntry, std::wstring>> entries;

    for (const auto& src : sources) {
        DWORD attr = GetFileAttributesW(src.c_str());
        if (attr == INVALID_FILE_ATTRIBUTES) continue;
        std::wstring nm = NameFromPath(src);
        ZipEntry ent;
        ent.nameUtf8 = WideToUtf8(nm);
        WIN32_FILE_ATTRIBUTE_DATA info{};
        GetFileAttributesExW(src.c_str(), GetFileExInfoStandard, &info);
        FiletimeToDos(info.ftLastWriteTime, ent.dosDate, ent.dosTime);
        if (attr & FILE_ATTRIBUTE_DIRECTORY) {
            ent.isDir = true;
            ent.nameUtf8 += "/";
            entries.push_back({ent, src});
            ZipCollectFiles(src, src, nm, entries);
        } else {
            ULARGE_INTEGER sz{};
            sz.LowPart = info.nFileSizeLow; sz.HighPart = info.nFileSizeHigh;
            ent.size = (uint32_t)std::min(sz.QuadPart, (ULONGLONG)0xFFFFFFFF);
            // CRC
            HANDLE hf = CreateFileW(src.c_str(), GENERIC_READ,
                FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
                OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (hf != INVALID_HANDLE_VALUE) {
                uint32_t crc = 0xFFFFFFFF;
                if (!Crc32Init) BuildCrc32();
                const DWORD BUF = 65536;
                std::vector<uint8_t> buf(BUF);
                DWORD rd;
                while (ReadFile(hf, buf.data(), BUF, &rd, nullptr) && rd > 0)
                    for (DWORD i = 0; i < rd; i++)
                        crc = Crc32Table[(crc ^ buf[i]) & 0xFF] ^ (crc >> 8);
                CloseHandle(hf);
                ent.crc32 = crc ^ 0xFFFFFFFF;
            }
            entries.push_back({ent, src});
        }
    }

    // Write local entries
    for (auto& [ent, path] : entries) {
        LARGE_INTEGER pos{};
        SetFilePointerEx(hOut, {}, &pos, FILE_CURRENT);
        ent.localOffset = (uint32_t)pos.QuadPart;
        if (!ZipWriteLocalEntry(hOut, ent, path)) {
            CloseHandle(hOut);
            DeleteFileW(zipPath.c_str());
            return false;
        }
    }

    // Central directory
    LARGE_INTEGER cdStart{};
    SetFilePointerEx(hOut, {}, &cdStart, FILE_CURRENT);
    for (auto& [ent, path] : entries) {
        std::vector<uint8_t> cd;
        cd.reserve(46 + ent.nameUtf8.size());
        cd.push_back(0x50); cd.push_back(0x4B); cd.push_back(0x01); cd.push_back(0x02);
        ZipPut16(cd, 20);                   // version made by
        ZipPut16(cd, 20);                   // version needed
        ZipPut16(cd, 0x0800);               // flags: UTF-8
        ZipPut16(cd, 0);                    // compression: store
        ZipPut16(cd, ent.dosTime);
        ZipPut16(cd, ent.dosDate);
        ZipPut32(cd, ent.crc32);
        ZipPut32(cd, ent.size);
        ZipPut32(cd, ent.size);
        ZipPut16(cd, (uint16_t)ent.nameUtf8.size());
        ZipPut16(cd, 0); ZipPut16(cd, 0);  // extra, comment
        ZipPut16(cd, 0); ZipPut16(cd, 0);  // disk start, int attr
        ZipPut32(cd, ent.isDir ? 0x10000000u : 0);
        ZipPut32(cd, ent.localOffset);
        for (char c : ent.nameUtf8) cd.push_back((uint8_t)c);
        DWORD wr;
        WriteFile(hOut, cd.data(), (DWORD)cd.size(), &wr, nullptr);
    }

    // End of central directory
    LARGE_INTEGER cdEnd{};
    SetFilePointerEx(hOut, {}, &cdEnd, FILE_CURRENT);
    uint32_t cdSize = (uint32_t)(cdEnd.QuadPart - cdStart.QuadPart);
    std::vector<uint8_t> eocd;
    eocd.push_back(0x50); eocd.push_back(0x4B); eocd.push_back(0x05); eocd.push_back(0x06);
    ZipPut16(eocd, 0); ZipPut16(eocd, 0);
    ZipPut16(eocd, (uint16_t)entries.size());
    ZipPut16(eocd, (uint16_t)entries.size());
    ZipPut32(eocd, cdSize);
    ZipPut32(eocd, (uint32_t)cdStart.QuadPart);
    ZipPut16(eocd, 0);
    DWORD wr;
    WriteFile(hOut, eocd.data(), (DWORD)eocd.size(), &wr, nullptr);
    CloseHandle(hOut);
    return true;
}

// =====================================================================
// ZIP EXTRACTION ENGINE
// =====================================================================

static bool ZipReadU16(const uint8_t* b, size_t off, size_t total, uint16_t& v) {
    if (off + 2 > total) return false;
    v = (uint16_t)(b[off] | (b[off+1] << 8)); return true;
}
static bool ZipReadU32(const uint8_t* b, size_t off, size_t total, uint32_t& v) {
    if (off + 4 > total) return false;
    v = (uint32_t)(b[off] | (b[off+1]<<8) | (b[off+2]<<16) | (b[off+3]<<24)); return true;
}

static bool ZipExtract(const std::wstring& zipPath, const std::wstring& destDir,
                       std::wstring& errOut)
{
    HANDLE hf = CreateFileW(zipPath.c_str(), GENERIC_READ,
        FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hf == INVALID_HANDLE_VALUE) { errOut = L"Cannot open ZIP file."; return false; }

    LARGE_INTEGER zipSz{};
    GetFileSizeEx(hf, &zipSz);
    if (zipSz.QuadPart < 22 || zipSz.QuadPart > 512*1024*1024LL) {
        CloseHandle(hf); errOut = L"ZIP too large or invalid."; return false;
    }

    // Read entire file into memory (for simplicity; max 512 MB checked above)
    size_t fsize = (size_t)zipSz.QuadPart;
    std::vector<uint8_t> data(fsize);
    DWORD rd; LARGE_INTEGER zero{};
    SetFilePointerEx(hf, zero, nullptr, FILE_BEGIN);
    if (!ReadFile(hf, data.data(), (DWORD)fsize, &rd, nullptr) || rd != (DWORD)fsize) {
        CloseHandle(hf); errOut = L"Failed to read ZIP."; return false;
    }
    CloseHandle(hf);

    // Find End of Central Directory record (search from end)
    int64_t eocdOff = -1;
    for (int64_t i = (int64_t)fsize - 22; i >= std::max(0LL, (int64_t)fsize - 65557); i--) {
        if (data[i]==0x50 && data[i+1]==0x4B && data[i+2]==0x05 && data[i+3]==0x06) {
            eocdOff = i; break;
        }
    }
    if (eocdOff < 0) { errOut = L"Not a valid ZIP file."; return false; }

    uint16_t numEntries; uint32_t cdSize, cdOffset;
    ZipReadU16(data.data(), eocdOff + 10, fsize, numEntries);
    ZipReadU32(data.data(), eocdOff + 12, fsize, cdSize);
    ZipReadU32(data.data(), eocdOff + 16, fsize, cdOffset);

    if (cdOffset >= fsize) { errOut = L"Corrupt ZIP: bad CD offset."; return false; }

    CreateDirectoryW(destDir.c_str(), nullptr);

    // Walk central directory
    size_t cdPos = cdOffset;
    for (uint16_t e = 0; e < numEntries; e++) {
        if (cdPos + 46 > fsize) break;
        if (data[cdPos]!=0x50||data[cdPos+1]!=0x4B||data[cdPos+2]!=0x01||data[cdPos+3]!=0x02) break;

        uint16_t compression, fnLen, extraLen, commentLen;
        uint32_t crc32, compSz, uncompSz, localOff;
        ZipReadU16(data.data(), cdPos+10, fsize, compression);
        ZipReadU32(data.data(), cdPos+16, fsize, crc32);
        ZipReadU32(data.data(), cdPos+20, fsize, compSz);
        ZipReadU32(data.data(), cdPos+24, fsize, uncompSz);
        ZipReadU16(data.data(), cdPos+28, fsize, fnLen);
        ZipReadU16(data.data(), cdPos+30, fsize, extraLen);
        ZipReadU16(data.data(), cdPos+32, fsize, commentLen);
        ZipReadU32(data.data(), cdPos+42, fsize, localOff);

        if (compression != 0) {
            // Skip compressed entries (deflate not implemented)
            cdPos += 46 + fnLen + extraLen + commentLen;
            continue;
        }

        std::string nameUtf8((char*)(data.data() + cdPos + 46), fnLen);
        cdPos += 46 + fnLen + extraLen + commentLen;

        // Convert name to wide (treat as UTF-8)
        int wn = MultiByteToWideChar(CP_UTF8, 0, nameUtf8.c_str(), (int)nameUtf8.size(), nullptr, 0);
        std::wstring namew(wn, L'\0');
        MultiByteToWideChar(CP_UTF8, 0, nameUtf8.c_str(), (int)nameUtf8.size(), &namew[0], wn);

        // Replace forward slashes
        for (auto& c : namew) if (c == L'/') c = L'\\';

        // Determine local header offset for actual data
        if (localOff + 30 > fsize) continue;
        uint16_t lfnLen, lExtraLen;
        ZipReadU16(data.data(), localOff + 26, fsize, lfnLen);
        ZipReadU16(data.data(), localOff + 28, fsize, lExtraLen);
        size_t dataStart = localOff + 30 + lfnLen + lExtraLen;
        if (dataStart > fsize) continue;

        std::wstring fullDest = destDir + L"\\" + namew;

        bool isDir = (!namew.empty() && namew.back() == L'\\') || uncompSz == 0 && compSz == 0;
        if (isDir) {
            // Strip trailing backslash if any
            std::wstring dirPath = fullDest;
            while (!dirPath.empty() && dirPath.back() == L'\\') dirPath.pop_back();
            // Create directory chain
            std::wstring partial = destDir;
            size_t p2 = dirPath.size();
            std::wstring rel = dirPath;
            if (rel.size() > destDir.size()) rel = rel.substr(destDir.size() + 1);
            size_t s = 0;
            while (s <= rel.size()) {
                size_t slash = rel.find(L'\\', s);
                if (slash == std::wstring::npos) slash = rel.size();
                partial = partial + L"\\" + rel.substr(s, slash - s);
                CreateDirectoryW(partial.c_str(), nullptr);
                s = slash + 1;
            }
            continue;
        }

        // Create parent directories
        std::wstring parent = fullDest;
        size_t lastSlash = parent.rfind(L'\\');
        if (lastSlash != std::wstring::npos) {
            parent = parent.substr(0, lastSlash);
            // Create chain
            std::wstring partial2 = destDir;
            if (parent.size() > destDir.size()) {
                std::wstring rel2 = parent.substr(destDir.size() + 1);
                size_t s2 = 0;
                while (s2 <= rel2.size()) {
                    size_t slash2 = rel2.find(L'\\', s2);
                    if (slash2 == std::wstring::npos) slash2 = rel2.size();
                    partial2 = partial2 + L"\\" + rel2.substr(s2, slash2 - s2);
                    CreateDirectoryW(partial2.c_str(), nullptr);
                    s2 = slash2 + 1;
                }
            }
        }

        if (dataStart + compSz > fsize) continue;

        HANDLE hOut = CreateFileW(fullDest.c_str(), GENERIC_WRITE, 0, nullptr,
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hOut == INVALID_HANDLE_VALUE) continue;

        DWORD wr;
        WriteFile(hOut, data.data() + dataStart, uncompSz, &wr, nullptr);
        CloseHandle(hOut);
    }
    return true;
}

// =====================================================================
// ZIP SELECTED (native)
// =====================================================================

static void DoZipSelected(Pane& p) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty() || p.tabs.empty()) return;
    std::wstring destDir = p.tabs[p.activeTab].path;
    std::wstring baseName = NameFromPath(sel[0]);
    std::wstring zipPath = CombinePath(destDir, baseName + L".zip");

    if (GetFileAttributesW(zipPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        int r = MessageBoxW(g.hwndMain, (L"\"" + zipPath + L"\" already exists. Overwrite?").c_str(),
            L"Compress", MB_YESNO | MB_ICONQUESTION);
        if (r != IDYES) return;
    }

    if (!ZipCreate(zipPath, sel)) {
        MessageBoxW(g.hwndMain, L"Failed to create ZIP archive.", L"Compress", MB_OK | MB_ICONERROR);
    }
    RefreshPane(p);
}

// =====================================================================
// EXTRACT ZIP (native)
// =====================================================================

static void DoExtractZip(Pane& p, const std::wstring& zipPath, bool chooseFolder) {
    std::wstring destDir;
    if (!p.tabs.empty()) destDir = p.tabs[p.activeTab].path;

    if (chooseFolder) {
        BROWSEINFOW bi{};
        bi.hwndOwner = g.hwndMain;
        bi.lpszTitle = L"Choose extraction folder:";
        bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
        PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
        if (!pidl) return;
        wchar_t buf[MAX_PATH]{};
        SHGetPathFromIDListW(pidl, buf);
        CoTaskMemFree(pidl);
        destDir = buf;
        if (destDir.empty()) return;
    }

    std::wstring err;
    if (!ZipExtract(zipPath, destDir, err)) {
        MessageBoxW(g.hwndMain, err.c_str(), L"Extract", MB_OK | MB_ICONERROR);
    } else {
        RefreshPane(p);
    }
}

// =====================================================================
// EXTRACT RAR (via WinRAR or 7-Zip CLI)
// =====================================================================

static std::wstring FindRarTool() {
    // Check WinRAR
    wchar_t pf[MAX_PATH]{}, pf86[MAX_PATH]{};
    GetEnvironmentVariableW(L"ProgramFiles", pf, MAX_PATH);
    GetEnvironmentVariableW(L"ProgramFiles(x86)", pf86, MAX_PATH);
    std::wstring candidates[] = {
        std::wstring(pf)   + L"\\WinRAR\\WinRAR.exe",
        std::wstring(pf86) + L"\\WinRAR\\WinRAR.exe",
        std::wstring(pf)   + L"\\WinRAR\\Rar.exe",
        std::wstring(pf86) + L"\\WinRAR\\Rar.exe",
        std::wstring(pf)   + L"\\7-Zip\\7z.exe",
        std::wstring(pf86) + L"\\7-Zip\\7z.exe",
    };
    for (auto& c : candidates) {
        if (GetFileAttributesW(c.c_str()) != INVALID_FILE_ATTRIBUTES) return c;
    }
    return L"";
}

static void DoExtractRar(Pane& p, const std::wstring& rarPath, bool chooseFolder) {
    std::wstring tool = FindRarTool();
    if (tool.empty()) {
        MessageBoxW(g.hwndMain,
            L"RAR extraction requires WinRAR or 7-Zip.\n"
            L"Please install one and try again.",
            L"Extract RAR", MB_OK | MB_ICONERROR);
        return;
    }

    std::wstring destDir;
    if (!p.tabs.empty()) destDir = p.tabs[p.activeTab].path;

    if (chooseFolder) {
        BROWSEINFOW bi{};
        bi.hwndOwner = g.hwndMain;
        bi.lpszTitle = L"Choose extraction folder:";
        bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
        PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
        if (!pidl) return;
        wchar_t buf[MAX_PATH]{};
        SHGetPathFromIDListW(pidl, buf);
        CoTaskMemFree(pidl);
        destDir = buf;
        if (destDir.empty()) return;
    }

    // Build command line: WinRAR uses "x -y archive dest\", 7-Zip uses "x archive -o dest -y"
    bool is7z = (tool.find(L"7z.exe") != std::wstring::npos);
    std::wstring rarDest = destDir;
    if (!rarDest.empty() && rarDest.back() != L'\\') rarDest += L'\\';
    std::wstring cmdLine;
    if (is7z) {
        cmdLine = L"\"" + tool + L"\" x \"" + rarPath + L"\" -o\"" + rarDest + L"\" -y";
    } else {
        cmdLine = L"\"" + tool + L"\" x -y \"" + rarPath + L"\" \"" + rarDest + L"\"";
    }

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION pi{};
    std::vector<wchar_t> cmdBuf(cmdLine.begin(), cmdLine.end());
    cmdBuf.push_back(0);

    if (CreateProcessW(nullptr, cmdBuf.data(), nullptr, nullptr, FALSE,
                       CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        WaitForSingleObject(pi.hProcess, INFINITE);
        DWORD exitCode = 0;
        GetExitCodeProcess(pi.hProcess, &exitCode);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        if (exitCode == 0 || exitCode == 1) { // 1 = warning (WinRAR), still ok
            RefreshPane(p);
        } else {
            MessageBoxW(g.hwndMain, L"Extraction failed.", L"Extract RAR", MB_OK | MB_ICONERROR);
        }
    } else {
        MessageBoxW(g.hwndMain, L"Failed to launch extraction tool.", L"Extract RAR", MB_OK | MB_ICONERROR);
    }
}

// =====================================================================
// SPLIT FILE
// =====================================================================

static void DoSplitFile(Pane& p) {
    auto sel = GetSelectedPaths(p);
    if (sel.size() != 1) { MessageBoxW(g.hwndMain, L"Select exactly one file to split.", L"Split", MB_OK|MB_ICONINFORMATION); return; }
    std::wstring src = sel[0];
    DWORD attr = GetFileAttributesW(src.c_str());
    if (attr == INVALID_FILE_ATTRIBUTES || (attr & FILE_ATTRIBUTE_DIRECTORY)) return;

    std::wstring mbStr = SimpleInputBox(g.hwndMain, L"Split File", L"Chunk size in MB:", L"10");
    if (mbStr.empty()) return;
    uint64_t chunkMB = (uint64_t)_wtoi64(mbStr.c_str());
    if (chunkMB < 1) chunkMB = 1;
    uint64_t chunkBytes = chunkMB * 1024 * 1024;

    HANDLE hin = CreateFileW(src.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
        OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
    if (hin == INVALID_HANDLE_VALUE) return;

    std::wstring dir = ParentPath(src);
    std::wstring base = NameFromPath(src);
    uint64_t partNum = 1;
    std::vector<uint8_t> buf((size_t)std::min(chunkBytes, (uint64_t)(4*1024*1024)));

    for (;;) {
        wchar_t partName[64];
        StringCchPrintfW(partName, 64, L".part%03llu", (unsigned long long)partNum);
        std::wstring outPath = CombinePath(dir, base + partName);
        HANDLE hout = CreateFileW(outPath.c_str(), GENERIC_WRITE, 0, nullptr,
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (hout == INVALID_HANDLE_VALUE) break;

        uint64_t written = 0;
        bool done = false;
        while (written < chunkBytes) {
            DWORD toRead = (DWORD)std::min((uint64_t)buf.size(), chunkBytes - written);
            DWORD got = 0;
            if (!ReadFile(hin, buf.data(), toRead, &got, nullptr) || got == 0) { done = true; break; }
            DWORD wr = 0;
            WriteFile(hout, buf.data(), got, &wr, nullptr);
            written += got;
        }
        CloseHandle(hout);
        partNum++;
        if (done) break;
    }
    CloseHandle(hin);
    RefreshPane(p);
}

// =====================================================================
// JOIN FILES
// =====================================================================

static void DoJoinFiles(Pane& p) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) { MessageBoxW(g.hwndMain, L"Select .partXXX files to join.", L"Join", MB_OK|MB_ICONINFORMATION); return; }
    std::sort(sel.begin(), sel.end(), [](const std::wstring& a, const std::wstring& b){ return _wcsicmp(a.c_str(), b.c_str()) < 0; });
    std::wstring first = sel[0];
    std::wstring base = first;
    size_t pos = base.rfind(L".part");
    if (pos != std::wstring::npos) base = base.substr(0, pos);
    std::wstring outPath = base + L"_joined";
    HANDLE hout = CreateFileW(outPath.c_str(), GENERIC_WRITE, 0, nullptr,
        CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hout == INVALID_HANDLE_VALUE) { MessageBoxW(g.hwndMain, L"Cannot create output file.", L"Join", MB_OK|MB_ICONERROR); return; }
    std::vector<uint8_t> buf(4*1024*1024);
    for (auto& part : sel) {
        HANDLE hin = CreateFileW(part.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
            OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
        if (hin == INVALID_HANDLE_VALUE) continue;
        DWORD got;
        while (ReadFile(hin, buf.data(), (DWORD)buf.size(), &got, nullptr) && got > 0) {
            DWORD wr;
            WriteFile(hout, buf.data(), got, &wr, nullptr);
        }
        CloseHandle(hin);
    }
    CloseHandle(hout);
    RefreshPane(p);
}

// =====================================================================
// FILE DIFF
// =====================================================================

static std::vector<std::wstring> ReadLines(const std::wstring& path) {
    std::vector<std::wstring> lines;
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, nullptr);
    if (h == INVALID_HANDLE_VALUE) return lines;
    LARGE_INTEGER sz{};
    GetFileSizeEx(h, &sz);
    if (sz.QuadPart > 4*1024*1024) { CloseHandle(h); return lines; }
    std::string raw((size_t)sz.QuadPart, '\0');
    DWORD got = 0;
    ReadFile(h, raw.data(), (DWORD)sz.QuadPart, &got, nullptr);
    CloseHandle(h);
    std::wstring all;
    int wn = MultiByteToWideChar(CP_UTF8, 0, raw.data(), got, nullptr, 0);
    if (wn > 0) { all.resize(wn); MultiByteToWideChar(CP_UTF8, 0, raw.data(), got, all.data(), wn); }
    size_t s = 0;
    for (size_t i = 0; i <= all.size(); i++) {
        if (i == all.size() || all[i] == L'\n') {
            std::wstring line = all.substr(s, i-s);
            if (!line.empty() && line.back()==L'\r') line.pop_back();
            lines.push_back(line);
            s = i+1;
        }
    }
    return lines;
}

static void DoFileDiff(Pane& p) {
    auto sel = GetSelectedPaths(p);
    std::wstring fileA, fileB;
    if (sel.size() >= 2) { fileA = sel[0]; fileB = sel[1]; }
    else if (sel.size() == 1) {
        // compare with same name in other pane
        int other = 1 - p.paneIndex;
        if (!g.panes[other].tabs.empty()) {
            fileB = CombinePath(g.panes[other].tabs[g.panes[other].activeTab].path, NameFromPath(sel[0]));
            fileA = sel[0];
        }
    }
    if (fileA.empty() || fileB.empty()) {
        MessageBoxW(g.hwndMain, L"Select 2 files or 1 file (will compare with same-name in other pane).", L"Diff", MB_OK|MB_ICONINFORMATION);
        return;
    }
    auto linesA = ReadLines(fileA);
    auto linesB = ReadLines(fileB);

    // Build diff output (simple line-by-line comparison)
    std::wstring result;
    result += L"--- " + NameFromPath(fileA) + L"\r\n";
    result += L"+++ " + NameFromPath(fileB) + L"\r\n\r\n";
    size_t maxLines = std::max(linesA.size(), linesB.size());
    for (size_t i = 0; i < maxLines; i++) {
        if (i < linesA.size() && i < linesB.size()) {
            if (linesA[i] == linesB[i]) result += L"  " + linesA[i] + L"\r\n";
            else { result += L"- " + linesA[i] + L"\r\n"; result += L"+ " + linesB[i] + L"\r\n"; }
        } else if (i < linesA.size()) result += L"- " + linesA[i] + L"\r\n";
        else result += L"+ " + linesB[i] + L"\r\n";
    }

    // Show in a simple scrollable dialog
    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME|WS_EX_TOPMOST, L"STATIC", L"File Diff",
        WS_POPUP|WS_CAPTION|WS_SYSMENU|WS_THICKFRAME, 0,0, 700,500, g.hwndMain, nullptr, g.hInst, nullptr);
    RECT rc; GetWindowRect(g.hwndMain, &rc);
    SetWindowPos(dlg, nullptr, rc.left+40, rc.top+40, 0,0, SWP_NOSIZE|SWP_NOZORDER);
    HWND ed = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", result.c_str(),
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|WS_HSCROLL|ES_MULTILINE|ES_READONLY|ES_AUTOHSCROLL|ES_AUTOVSCROLL,
        4,4, 688,456, dlg, (HMENU)ID_DIFF_LIST, g.hInst, nullptr);
    SendMessageW(ed, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    HWND closeBtn = CreateWindowExW(0, L"BUTTON", L"Close", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
        300,464, 100,26, dlg, (HMENU)ID_DIFF_CANCEL, g.hInst, nullptr);
    SendMessageW(closeBtn, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    struct DiffDlg {
        static LRESULT CALLBACK Proc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
            if (m == WM_COMMAND && LOWORD(wp) == ID_DIFF_CANCEL) DestroyWindow(h);
            if (m == WM_CLOSE) DestroyWindow(h);
            return DefWindowProcW(h, m, wp, lp);
        }
    };
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)DiffDlg::Proc);
    ShowWindow(dlg, SW_SHOW);
}

// =====================================================================
// FOLDER DIFF
// =====================================================================

static void DoFolderDiff(Pane& p) {
    std::wstring dirA, dirB;
    auto sel = GetSelectedPaths(p);
    if (sel.size() >= 2) { dirA = sel[0]; dirB = sel[1]; }
    else {
        dirA = p.tabs.empty() ? L"" : p.tabs[p.activeTab].path;
        int other = 1 - p.paneIndex;
        dirB = g.panes[other].tabs.empty() ? L"" : g.panes[other].tabs[g.panes[other].activeTab].path;
    }
    if (dirA.empty() || dirB.empty()) return;

    auto listDir = [](const std::wstring& dir) {
        std::unordered_map<std::wstring, FileEntry> m;
        std::wstring pat = dir + L"\\*";
        WIN32_FIND_DATAW fd;
        HANDLE h = FindFirstFileW(pat.c_str(), &fd);
        if (h == INVALID_HANDLE_VALUE) return m;
        do {
            if (wcscmp(fd.cFileName,L".")==0||wcscmp(fd.cFileName,L"..")==0) continue;
            FileEntry e;
            e.name = fd.cFileName;
            e.size = ((uint64_t)fd.nFileSizeHigh<<32)|fd.nFileSizeLow;
            e.modified = fd.ftLastWriteTime;
            e.attr = fd.dwFileAttributes;
            std::wstring key = e.name;
            for (auto& c : key) c = (wchar_t)towupper(c);
            m[key] = e;
        } while (FindNextFileW(h,&fd));
        FindClose(h);
        return m;
    };

    auto mapA = listDir(dirA);
    auto mapB = listDir(dirB);

    std::wstring result;
    result += L"Comparing:\r\n  A: " + dirA + L"\r\n  B: " + dirB + L"\r\n\r\n";
    for (auto& kv : mapA) {
        auto it = mapB.find(kv.first);
        if (it == mapB.end()) result += L"Only in A: " + kv.second.name + L"\r\n";
        else if (kv.second.size != it->second.size ||
                 CompareFileTime(&kv.second.modified, &it->second.modified) != 0)
            result += L"Modified:  " + kv.second.name + L"\r\n";
    }
    for (auto& kv : mapB) {
        if (mapA.find(kv.first) == mapA.end())
            result += L"Only in B: " + kv.second.name + L"\r\n";
    }
    if (result.find(L"Only") == std::wstring::npos && result.find(L"Modified") == std::wstring::npos)
        result += L"(No differences)\r\n";

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME|WS_EX_TOPMOST, L"STATIC", L"Folder Diff",
        WS_POPUP|WS_CAPTION|WS_SYSMENU|WS_THICKFRAME, 0,0, 600,400, g.hwndMain, nullptr, g.hInst, nullptr);
    RECT rc; GetWindowRect(g.hwndMain, &rc);
    SetWindowPos(dlg, nullptr, rc.left+50, rc.top+50, 0,0, SWP_NOSIZE|SWP_NOZORDER);
    HWND ed = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", result.c_str(),
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|ES_MULTILINE|ES_READONLY|ES_AUTOVSCROLL,
        4,4, 588,356, dlg, nullptr, g.hInst, nullptr);
    SendMessageW(ed, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    HWND closeBtn = CreateWindowExW(0, L"BUTTON", L"Close", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
        250,368, 100,26, dlg, (HMENU)ID_DIFF_CANCEL, g.hInst, nullptr);
    SendMessageW(closeBtn, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    struct FDlg { static LRESULT CALLBACK P(HWND h,UINT m,WPARAM wp,LPARAM l) {
        if (m==WM_COMMAND&&LOWORD(wp)==ID_DIFF_CANCEL) DestroyWindow(h);
        if (m==WM_CLOSE) DestroyWindow(h);
        return DefWindowProcW(h,m,wp,l);
    }};
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)FDlg::P);
    ShowWindow(dlg, SW_SHOW);
}

// =====================================================================
// DUPLICATE FINDER
// =====================================================================

static void DoFindDuplicates(Pane& p) {
    if (g.dupeRunning.load()) { MessageBoxW(g.hwndMain, L"Scan already in progress.", L"Dupes", MB_OK); return; }
    auto sel = GetSelectedPaths(p);
    std::wstring root = (sel.size()==1 && DirExists(sel[0])) ? sel[0] : (p.tabs.empty() ? L"" : p.tabs[p.activeTab].path);
    if (root.empty()) return;

    g.dupeRunning.store(true);
    g.dupeGroups.clear();
    if (g.dupeThread.joinable()) g.dupeThread.join();
    g.dupeThread = std::thread([root](){
        // Collect all files
        std::vector<std::wstring> allFiles;
        std::function<void(const std::wstring&)> collect = [&](const std::wstring& dir) {
            std::wstring pat = dir + L"\\*";
            WIN32_FIND_DATAW fd;
            HANDLE h = FindFirstFileW(pat.c_str(), &fd);
            if (h == INVALID_HANDLE_VALUE) return;
            do {
                if (wcscmp(fd.cFileName,L".")==0||wcscmp(fd.cFileName,L"..")==0) continue;
                std::wstring full = dir + L"\\" + fd.cFileName;
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) collect(full);
                else { uint64_t sz = ((uint64_t)fd.nFileSizeHigh<<32)|fd.nFileSizeLow; if (sz > 0) allFiles.push_back(full); }
            } while (FindNextFileW(h,&fd));
            FindClose(h);
        };
        collect(root);

        // Group by size
        std::unordered_map<uint64_t, std::vector<std::wstring>> bySz;
        for (auto& f : allFiles) {
            HANDLE h = CreateFileW(f.c_str(), 0, FILE_SHARE_READ|FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (h == INVALID_HANDLE_VALUE) continue;
            LARGE_INTEGER sz{}; GetFileSizeEx(h, &sz); CloseHandle(h);
            if (sz.QuadPart > 0) bySz[(uint64_t)sz.QuadPart].push_back(f);
        }

        std::vector<std::vector<std::wstring>> groups;
        for (auto& kv : bySz) {
            if (kv.second.size() < 2) continue;
            groups.push_back(kv.second);
        }

        g.dupeGroups = groups;
        g.dupeRunning.store(false);
        if (g.hwndMain) PostMessageW(g.hwndMain, WM_DUPE_READY, 0, 0);
    });
    g.dupeThread.detach();
    SetWindowTextW(g.hwndStatus, L"Scanning for duplicates...");
}

static void ShowDupeResults() {
    if (g.dupeGroups.empty()) {
        MessageBoxW(g.hwndMain, L"No duplicate files found.", L"Duplicates", MB_OK|MB_ICONINFORMATION);
        return;
    }

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME|WS_EX_TOPMOST, L"STATIC", L"Duplicate Files",
        WS_POPUP|WS_CAPTION|WS_SYSMENU|WS_THICKFRAME, 0,0, 650,450, g.hwndMain, nullptr, g.hInst, nullptr);
    RECT rc; GetWindowRect(g.hwndMain, &rc);
    SetWindowPos(dlg, nullptr, rc.left+30, rc.top+30, 0,0, SWP_NOSIZE|SWP_NOZORDER);

    HWND lb = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_HASSTRINGS|LBS_NOINTEGRALHEIGHT,
        4,4, 638,370, dlg, (HMENU)ID_DUPE_LIST, g.hInst, nullptr);
    SendMessageW(lb, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    for (size_t gi = 0; gi < g.dupeGroups.size(); gi++) {
        wchar_t hdr[64];
        StringCchPrintfW(hdr, 64, L"--- Group %zu (%zu files) ---", gi+1, g.dupeGroups[gi].size());
        SendMessageW(lb, LB_ADDSTRING, 0, (LPARAM)hdr);
        for (auto& f : g.dupeGroups[gi])
            SendMessageW(lb, LB_ADDSTRING, 0, (LPARAM)f.c_str());
    }

    HWND closeBtn = CreateWindowExW(0, L"BUTTON", L"Close", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
        275,420, 100,26, dlg, (HMENU)ID_DUPE_CANCEL, g.hInst, nullptr);
    SendMessageW(closeBtn, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    struct DupeDlg { static LRESULT CALLBACK P(HWND h,UINT m,WPARAM wp,LPARAM l) {
        if (m==WM_COMMAND&&LOWORD(wp)==ID_DUPE_CANCEL) DestroyWindow(h);
        if (m==WM_CLOSE) DestroyWindow(h);
        return DefWindowProcW(h,m,wp,l);
    }};
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)DupeDlg::P);
    ShowWindow(dlg, SW_SHOW);
}

// =====================================================================
// PROCESS LOCK INSPECTOR
// =====================================================================

static void ShowFileLockInfo(const std::wstring& path) {
    // Try to open the file exclusively
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ|GENERIC_WRITE, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h != INVALID_HANDLE_VALUE) {
        CloseHandle(h);
        MessageBoxW(g.hwndMain, L"File is not currently locked.", L"Lock Info", MB_OK|MB_ICONINFORMATION);
        return;
    }
    if (GetLastError() == ERROR_SHARING_VIOLATION) {
        int r = MessageBoxW(g.hwndMain,
            L"File is locked by another process.\n\nWould you like to open Task Manager to find the locking process?",
            L"File Locked", MB_YESNO|MB_ICONWARNING);
        if (r == IDYES)
            ShellExecuteW(nullptr, L"open", L"taskmgr.exe", nullptr, nullptr, SW_SHOW);
    } else {
        MessageBoxW(g.hwndMain, L"Cannot access file.", L"Lock Info", MB_OK|MB_ICONWARNING);
    }
}

// =====================================================================
// GIT STATUS
// =====================================================================

static void StartGitStatus(const std::wstring& dir) {
    if (g.gitRunning.exchange(true)) return;
    if (g.gitThread.joinable()) g.gitThread.detach();
    g.gitThread = std::thread([dir](){
        // Find git root by walking up
        std::wstring root = dir;
        bool found = false;
        for (int i = 0; i < 10; i++) {
            if (GetFileAttributesW((root + L"\\.git").c_str()) != INVALID_FILE_ATTRIBUTES) { found = true; break; }
            std::wstring parent = ParentPath(root);
            if (parent == root) break;
            root = parent;
        }
        if (!found) { g.gitRunning.store(false); return; }

        std::wstring cmd = L"git -C \"" + root + L"\" status --porcelain 2>nul";
        FILE* f = _wpopen(cmd.c_str(), L"rt");
        if (!f) { g.gitRunning.store(false); return; }

        std::unordered_map<std::wstring, wchar_t> newStatus;
        wchar_t buf[512];
        while (fgetws(buf, 512, f)) {
            if (wcslen(buf) < 4) continue;
            wchar_t st = (buf[0] != L' ' && buf[0] != L'?') ? buf[0] : buf[1];
            if (buf[0] == L'?' && buf[1] == L'?') st = L'?';
            std::wstring fname = std::wstring(buf + 3);
            if (!fname.empty() && (fname.back()==L'\n'||fname.back()==L'\r')) fname.pop_back();
            if (!fname.empty() && (fname.back()==L'\n'||fname.back()==L'\r')) fname.pop_back();
            // get top-level component
            size_t sl = fname.find_first_of(L"/\\");
            if (sl != std::wstring::npos) fname = fname.substr(0, sl);
            if (!fname.empty()) newStatus[fname] = st;
        }
        _pclose(f);
        {
            std::lock_guard<std::mutex> lk(g.iconMu);
            g.gitStatus = newStatus;
        }
        g.gitRunning.store(false);
        for (int i = 0; i < 2; i++)
            if (g.panes[i].hwndList) InvalidateRect(g.panes[i].hwndList, nullptr, FALSE);
    });
    g.gitThread.detach();
}

// =====================================================================
// AUTO-REFRESH (ReadDirectoryChangesW via FindFirstChangeNotification)
// =====================================================================

static void StartWatchThread(int paneIdx) {
    if (paneIdx < 0 || paneIdx > 1) return;
    g.watchStop.store(false);
    if (g.watchThreads[paneIdx].joinable()) {
        g.watchStop.store(true);
        // Can't easily join a detached thread; just overwrite
        g.watchStop.store(false);
    }
    if (g.panes[paneIdx].tabs.empty()) return;
    std::wstring path = g.panes[paneIdx].tabs[g.panes[paneIdx].activeTab].path;
    int pidx = paneIdx;
    g.watchThreads[paneIdx] = std::thread([pidx, path](){
        HANDLE h = FindFirstChangeNotificationW(path.c_str(), FALSE,
            FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME |
            FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE);
        if (h == INVALID_HANDLE_VALUE) return;
        while (!g.watchStop.load()) {
            DWORD r = WaitForSingleObject(h, 400);
            if (g.watchStop.load()) break;
            if (r == WAIT_OBJECT_0) {
                if (g.panes[pidx].hwnd && g.settingAutoRefresh)
                    PostMessageW(g.panes[pidx].hwnd, WM_AUTOREFRESH, (WPARAM)pidx, 0);
                FindNextChangeNotification(h);
            }
        }
        FindCloseChangeNotification(h);
    });
    g.watchThreads[paneIdx].detach();
}

static void StopWatchThreads() {
    g.watchStop.store(true);
}

// =====================================================================
// CLIPBOARD HISTORY
// =====================================================================

static void ShowClipHistoryDialog(Pane& p) {
    if (g.clipHistory.empty()) {
        MessageBoxW(g.hwndMain, L"Clipboard history is empty.", L"Clip History", MB_OK|MB_ICONINFORMATION);
        return;
    }

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME|WS_EX_TOPMOST, L"STATIC", L"Clipboard History",
        WS_POPUP|WS_CAPTION|WS_SYSMENU, 0,0, 500,300, g.hwndMain, nullptr, g.hInst, nullptr);
    RECT rc; GetWindowRect(g.hwndMain, &rc);
    SetWindowPos(dlg, nullptr, rc.left+(rc.right-rc.left-500)/2, rc.top+(rc.bottom-rc.top-300)/2, 0,0, SWP_NOSIZE|SWP_NOZORDER);

    HWND lb = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_HASSTRINGS|LBS_NOINTEGRALHEIGHT,
        4,4, 488,220, dlg, (HMENU)ID_CP_LIST, g.hInst, nullptr);
    SendMessageW(lb, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    for (size_t i = 0; i < g.clipHistory.size(); i++) {
        auto& ce = g.clipHistory[i];
        std::wstring line = (ce.cut ? L"[CUT] " : L"[COPY] ");
        if (!ce.files.empty()) line += NameFromPath(ce.files[0]);
        if (ce.files.size() > 1) line += L" (+" + std::to_wstring(ce.files.size()-1) + L" more)";
        SendMessageW(lb, LB_ADDSTRING, 0, (LPARAM)line.c_str());
    }

    HWND useBtn = CreateWindowExW(0, L"BUTTON", L"Paste Selected", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 4,232, 130,26, dlg, (HMENU)ID_CP_OK, g.hInst, nullptr);
    SendMessageW(useBtn, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    HWND closeBtn = CreateWindowExW(0, L"BUTTON", L"Close", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, 390,232, 100,26, dlg, (HMENU)ID_CP_CANCEL, g.hInst, nullptr);
    SendMessageW(closeBtn, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    struct ClipDlg {
        static LRESULT CALLBACK P(HWND h, UINT m, WPARAM wp, LPARAM lp) {
            if (m == WM_COMMAND) {
                if (LOWORD(wp) == ID_CP_OK) {
                    HWND lb2 = GetDlgItem(h, ID_CP_LIST);
                    int sel = (int)SendMessageW(lb2, LB_GETCURSEL, 0, 0);
                    if (sel >= 0 && sel < (int)g.clipHistory.size()) {
                        auto& ce = g.clipHistory[sel];
                        g.clipboardFiles = ce.files;
                        g.clipboardCut = ce.cut;
                        g.clipboardCutSet.clear();
                        if (ce.cut) for (auto& f : ce.files) g.clipboardCutSet.insert(f);
                        DoPaste(g.panes[g.activePane]);
                    }
                    DestroyWindow(h);
                } else if (LOWORD(wp) == ID_CP_CANCEL) DestroyWindow(h);
            }
            if (m == WM_CLOSE) DestroyWindow(h);
            return DefWindowProcW(h, m, wp, lp);
        }
    };
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)ClipDlg::P);
    ShowWindow(dlg, SW_SHOW);
    (void)p;
}

// =====================================================================
// QUICK OPEN (fuzzy recursive file search)
// =====================================================================

static std::mutex g_qoMutex;
static std::vector<std::wstring> g_qoResults;
static std::atomic<bool> g_qoStop{false};
static HWND g_qoListbox = nullptr;
static HWND g_qoDialog = nullptr;
static Pane* g_qoPane = nullptr;

static void QOSearch(const std::wstring& dir, const std::wstring& query, int depth) {
    if (g_qoStop.load() || depth > 5) return;
    std::wstring pat = dir + L"\\*";
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pat.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return;
    do {
        if (g_qoStop.load()) break;
        if (wcscmp(fd.cFileName,L".")==0||wcscmp(fd.cFileName,L"..")==0) continue;
        std::wstring full = dir + L"\\" + fd.cFileName;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            QOSearch(full, query, depth+1);
        } else {
            std::wstring name = fd.cFileName;
            std::wstring uname = name, uq = query;
            for (auto& c : uname) c = (wchar_t)towupper(c);
            for (auto& c : uq) c = (wchar_t)towupper(c);
            if (uname.find(uq) != std::wstring::npos) {
                std::lock_guard<std::mutex> lk(g_qoMutex);
                g_qoResults.push_back(full);
                if (g_qoDialog && g_qoListbox)
                    PostMessageW(g_qoDialog, WM_USER+20, 0, 0);
            }
        }
    } while (FindNextFileW(h,&fd));
    FindClose(h);
}

static LRESULT CALLBACK QuickOpenProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_COMMAND: {
            int id = LOWORD(w), code = HIWORD(w);
            if (id == ID_QO_EDIT && code == EN_CHANGE) {
                wchar_t buf[256] = {};
                GetWindowTextW(GetDlgItem(hwnd, ID_QO_EDIT), buf, 255);
                std::wstring query = buf;
                // Reset and restart search
                g_qoStop.store(true);
                {
                    std::lock_guard<std::mutex> lk(g_qoMutex);
                    g_qoResults.clear();
                }
                SendMessageW(GetDlgItem(hwnd, ID_QO_LIST), LB_RESETCONTENT, 0, 0);
                if (query.size() < 2) { g_qoStop.store(false); return 0; }
                g_qoStop.store(false);
                std::wstring searchRoot = g_qoPane && !g_qoPane->tabs.empty() ?
                    g_qoPane->tabs[g_qoPane->activeTab].path : L"C:\\";
                std::thread([query, searchRoot](){
                    QOSearch(searchRoot, query, 0);
                }).detach();
            }
            if (id == ID_QO_LIST && code == LBN_DBLCLK) {
                int sel = (int)SendMessageW(GetDlgItem(hwnd, ID_QO_LIST), LB_GETCURSEL, 0, 0);
                std::wstring path;
                {
                    std::lock_guard<std::mutex> lk(g_qoMutex);
                    if (sel >= 0 && sel < (int)g_qoResults.size()) path = g_qoResults[sel];
                }
                if (!path.empty() && g_qoPane) {
                    g_qoStop.store(true);
                    OpenPath(*g_qoPane, path);
                    DestroyWindow(hwnd);
                }
            }
            if (id == ID_QO_CANCEL) { g_qoStop.store(true); DestroyWindow(hwnd); }
            return 0;
        }
        case WM_USER+20: {
            // Update listbox from results
            HWND lb = GetDlgItem(hwnd, ID_QO_LIST);
            SendMessageW(lb, LB_RESETCONTENT, 0, 0);
            std::lock_guard<std::mutex> lk(g_qoMutex);
            for (auto& r : g_qoResults)
                SendMessageW(lb, LB_ADDSTRING, 0, (LPARAM)r.c_str());
            return 0;
        }
        case WM_KEYDOWN:
            if (w == VK_ESCAPE) { g_qoStop.store(true); DestroyWindow(hwnd); return 0; }
            break;
        case WM_DESTROY:
            g_qoStop.store(true);
            g.hwndQuickOpen = nullptr;
            g_qoDialog = nullptr;
            return 0;
        case WM_CLOSE:
            g_qoStop.store(true);
            DestroyWindow(hwnd);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, w, l);
}

static void ShowQuickOpen(Pane& p) {
    if (g.hwndQuickOpen) { SetForegroundWindow(g.hwndQuickOpen); return; }
    g_qoPane = &p;
    g_qoResults.clear();
    g_qoStop.store(false);

    RECT main; GetWindowRect(g.hwndMain, &main);
    int W = 500, H = 350;
    int X = main.left + (main.right-main.left-W)/2;
    int Y = main.top + 60;

    HWND dlg = CreateWindowExW(WS_EX_TOPMOST|WS_EX_DLGMODALFRAME, CLASS_QUICKOPEN, L"Quick Open (Ctrl+K)",
        WS_POPUP|WS_CAPTION|WS_SYSMENU, X,Y,W,H, g.hwndMain, nullptr, g.hInst, nullptr);
    if (!dlg) return;

    HWND ed = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL, 4,4, W-12, 26, dlg, (HMENU)ID_QO_EDIT, g.hInst, nullptr);
    SendMessageW(ed, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    g_qoListbox = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_HASSTRINGS|LBS_NOINTEGRALHEIGHT,
        4,34, W-12, H-80, dlg, (HMENU)ID_QO_LIST, g.hInst, nullptr);
    SendMessageW(g_qoListbox, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    HWND closeBtn = CreateWindowExW(0, L"BUTTON", L"Close", WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON,
        W/2-50, H-40, 100, 26, dlg, (HMENU)ID_QO_CANCEL, g.hInst, nullptr);
    SendMessageW(closeBtn, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    g.hwndQuickOpen = dlg;
    g_qoDialog = dlg;
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)QuickOpenProc);
    ShowWindow(dlg, SW_SHOW);
    SetFocus(ed);
}

// =====================================================================
// COMMAND PALETTE
// =====================================================================

struct CmdEntry { const wchar_t* name; int id; };
static const CmdEntry s_commands[] = {
    {L"Open",                 IDM_OPEN},
    {L"Open in New Tab",      IDM_OPEN_NEW_TAB},
    {L"Open in Other Pane",   IDM_OPEN_OTHER_PANE},
    {L"Copy",                 IDM_COPY},
    {L"Cut",                  IDM_CUT},
    {L"Paste",                IDM_PASTE},
    {L"Delete (Recycle Bin)", IDM_DELETE},
    {L"Delete Permanently",   IDM_DELETE_PERM},
    {L"Rename",               IDM_RENAME},
    {L"Bulk Rename",          IDM_BULK_RENAME},
    {L"New Folder",           IDM_NEW_FOLDER},
    {L"Refresh",              IDM_REFRESH},
    {L"Properties",           IDM_PROPERTIES},
    {L"Sort by Name",         IDM_SORT_NAME},
    {L"Sort by Size",         IDM_SORT_SIZE},
    {L"Sort by Date",         IDM_SORT_DATE},
    {L"Sort by Type",         IDM_SORT_TYPE},
    {L"Toggle Hidden Files",  IDM_TOGGLE_HIDDEN},
    {L"Select by Pattern",    IDM_SELECT_PATTERN},
    {L"Invert Selection",     IDM_INVERT_SEL},
    {L"Copy Full Path",       IDM_COPY_PATH_FULL},
    {L"Copy Filename",        IDM_COPY_PATH_NAME},
    {L"Copy Directory",       IDM_COPY_PATH_DIR},
    {L"Open Terminal Here",   IDM_OPEN_TERMINAL},
    {L"Zip Selected",         IDM_ZIP_SELECTED},
    {L"Split File",           IDM_SPLIT_FILE},
    {L"Join Files",           IDM_JOIN_FILES},
    {L"File Diff",            IDM_FILE_DIFF},
    {L"Folder Diff",          IDM_FOLDER_DIFF},
    {L"Find Duplicates",      IDM_FIND_DUPES},
    {L"File Lock Info",       IDM_PROC_LOCK},
    {L"Toggle Grid View",     IDM_GRID_VIEW},
    {L"Toggle Hex View",      IDM_HEX_VIEW},
    {L"Undo",                 IDM_UNDO},
    {L"Settings",             IDM_SETTINGS},
    {L"Clipboard History",    IDM_CLIP_HISTORY},
    {L"Quick Open",           IDM_QUICK_OPEN},
    {L"Connect to Server",    IDM_CONNECT_SERVER},
    {L"Git Status Refresh",   IDM_GIT_STATUS},
    {L"Toggle Dual Pane",     IDM_COMPUTE_HASH},
    {L"Folder Size",          IDM_FOLDER_SIZE},
    {L"Pin to Sidebar",       IDM_PIN_SIDEBAR},
};
static const int s_commandCount = (int)(sizeof(s_commands)/sizeof(s_commands[0]));

static std::vector<int> g_palFiltered;
static HWND g_palDialog = nullptr;
static Pane* g_palPane = nullptr;

static void PalFilter(HWND dlg, const std::wstring& query) {
    HWND lb = GetDlgItem(dlg, ID_PAL_LIST);
    SendMessageW(lb, LB_RESETCONTENT, 0, 0);
    g_palFiltered.clear();
    for (int i = 0; i < s_commandCount; i++) {
        std::wstring name = s_commands[i].name;
        std::wstring uname = name, uq = query;
        for (auto& c : uname) c = (wchar_t)towupper(c);
        for (auto& c : uq) c = (wchar_t)towupper(c);
        if (uq.empty() || uname.find(uq) != std::wstring::npos) {
            SendMessageW(lb, LB_ADDSTRING, 0, (LPARAM)name.c_str());
            g_palFiltered.push_back(i);
        }
    }
    if (!g_palFiltered.empty()) SendMessageW(lb, LB_SETCURSEL, 0, 0);
}

static void PalExecute(HWND dlg) {
    HWND lb = GetDlgItem(dlg, ID_PAL_LIST);
    int sel = (int)SendMessageW(lb, LB_GETCURSEL, 0, 0);
    if (sel < 0 || sel >= (int)g_palFiltered.size()) return;
    int id = s_commands[g_palFiltered[sel]].id;
    DestroyWindow(dlg);
    if (g_palPane)
        PostMessageW(g_palPane->hwnd, WM_COMMAND, (WPARAM)id, 0);
}

static LRESULT CALLBACK CmdPalProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_COMMAND: {
            int id = LOWORD(w), code = HIWORD(w);
            if (id == ID_PAL_EDIT && code == EN_CHANGE) {
                wchar_t buf[256] = {};
                GetWindowTextW(GetDlgItem(hwnd, ID_PAL_EDIT), buf, 255);
                PalFilter(hwnd, buf);
            }
            if (id == ID_PAL_OK && code == BN_CLICKED) PalExecute(hwnd);
            if (id == ID_PAL_CANCEL) { g.hwndCmdPal = nullptr; DestroyWindow(hwnd); }
            if (id == ID_PAL_LIST && code == LBN_DBLCLK) PalExecute(hwnd);
            return 0;
        }
        case WM_KEYDOWN:
            if (w == VK_ESCAPE) { g.hwndCmdPal = nullptr; DestroyWindow(hwnd); return 0; }
            if (w == VK_RETURN) { PalExecute(hwnd); return 0; }
            if (w == VK_DOWN) { HWND lb = GetDlgItem(hwnd, ID_PAL_LIST); int s=(int)SendMessageW(lb,LB_GETCURSEL,0,0); SendMessageW(lb,LB_SETCURSEL,s+1,0); return 0; }
            if (w == VK_UP)   { HWND lb = GetDlgItem(hwnd, ID_PAL_LIST); int s=(int)SendMessageW(lb,LB_GETCURSEL,0,0); if(s>0) SendMessageW(lb,LB_SETCURSEL,s-1,0); return 0; }
            break;
        case WM_DESTROY:
            g.hwndCmdPal = nullptr;
            g_palDialog = nullptr;
            return 0;
        case WM_CLOSE:
            g.hwndCmdPal = nullptr;
            DestroyWindow(hwnd);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, w, l);
}

static void ShowCmdPalette(Pane& p) {
    if (g.hwndCmdPal) { SetForegroundWindow(g.hwndCmdPal); return; }
    g_palPane = &p;
    g_palFiltered.clear();

    RECT main; GetWindowRect(g.hwndMain, &main);
    int W = 480, H = 380;
    int X = main.left + (main.right-main.left-W)/2;
    int Y = main.top + 50;

    HWND dlg = CreateWindowExW(WS_EX_TOPMOST|WS_EX_DLGMODALFRAME, CLASS_CMDPAL, L"Command Palette (Ctrl+P)",
        WS_POPUP|WS_CAPTION|WS_SYSMENU, X,Y,W,H, g.hwndMain, nullptr, g.hInst, nullptr);
    if (!dlg) return;

    // Search edit
    HWND ed = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD|WS_VISIBLE|ES_AUTOHSCROLL, 4,4, W-12, 26, dlg, (HMENU)(INT_PTR)(ID_PAL_EDIT), g.hInst, nullptr);
    SendMessageW(ed, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    HWND lb = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD|WS_VISIBLE|WS_VSCROLL|LBS_HASSTRINGS|LBS_NOINTEGRALHEIGHT,
        4,34, W-12, H-80, dlg, (HMENU)ID_PAL_LIST, g.hInst, nullptr);
    SendMessageW(lb, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    HWND ok  = CreateWindowExW(0,L"BUTTON",L"Run",   WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, W/2-105,H-38,100,26,dlg,(HMENU)ID_PAL_OK,g.hInst,nullptr);
    HWND can = CreateWindowExW(0,L"BUTTON",L"Cancel",WS_CHILD|WS_VISIBLE|BS_PUSHBUTTON, W/2+5,  H-38,100,26,dlg,(HMENU)ID_PAL_CANCEL,g.hInst,nullptr);
    SendMessageW(ok,  WM_SETFONT, (WPARAM)g.hFont, TRUE);
    SendMessageW(can, WM_SETFONT, (WPARAM)g.hFont, TRUE);

    PalFilter(dlg, L"");

    g.hwndCmdPal = dlg;
    g_palDialog = dlg;
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)CmdPalProc);
    PalFilter(dlg, L"");
    ShowWindow(dlg, SW_SHOW);
    SetFocus(ed);
}

// =====================================================================
// GRID VIEW TOGGLE
// =====================================================================

static void ToggleGridView(Pane& p) {
    g.viewMode[p.paneIndex] ^= 1;
    if (g.viewMode[p.paneIndex] == 1) {
        // Switch to icon view
        DWORD style = (DWORD)GetWindowLongW(p.hwndList, GWL_STYLE);
        style &= ~LVS_TYPEMASK;
        style |= LVS_ICON;
        SetWindowLongW(p.hwndList, GWL_STYLE, style);
        ListView_SetIconSpacing(p.hwndList, 100, 80);
        if (g.sysImageList) {
            HIMAGELIST lil = nullptr;
            SHFILEINFOW sfi{};
            lil = (HIMAGELIST)SHGetFileInfoW(L"C:\\", 0, &sfi, sizeof(sfi),
                SHGFI_SYSICONINDEX | SHGFI_LARGEICON);
            if (lil) ListView_SetImageList(p.hwndList, lil, LVSIL_NORMAL);
        }
    } else {
        // Switch back to details
        DWORD style = (DWORD)GetWindowLongW(p.hwndList, GWL_STYLE);
        style &= ~LVS_TYPEMASK;
        style |= LVS_REPORT;
        SetWindowLongW(p.hwndList, GWL_STYLE, style);
        if (g.sysImageList) ListView_SetImageList(p.hwndList, g.sysImageList, LVSIL_SMALL);
    }
    InvalidateRect(p.hwndList, nullptr, TRUE);
}

// =====================================================================
// HEX VIEW TOGGLE
// =====================================================================

static void ToggleHexView(Pane& p) {
    g.hexView[p.paneIndex] ^= 1;
    if (g.hwndPreview) InvalidateRect(g.hwndPreview, nullptr, FALSE);
}

// =====================================================================
// COLOR TAGS
// =====================================================================

static void SetColorTag(Pane& p) {
    auto sel = GetSelectedPaths(p);
    if (sel.empty()) return;
    static const struct { COLORREF c; const wchar_t* name; } colors[] = {
        {RGB(200,50,50),  L"Red"},
        {RGB(50,180,50),  L"Green"},
        {RGB(50,120,220), L"Blue"},
        {RGB(220,180,50), L"Yellow"},
        {RGB(180,80,180), L"Purple"},
        {RGB(50,190,190), L"Cyan"},
        {RGB(220,130,50), L"Orange"},
        {0,               L"None (remove)"},
    };
    HMENU m = CreatePopupMenu();
    for (int i = 0; i < 8; i++) AppendMenuW(m, MF_STRING, i+1, colors[i].name);
    POINT pt; GetCursorPos(&pt);
    int cmd = TrackPopupMenu(m, TPM_RETURNCMD|TPM_RIGHTBUTTON, pt.x, pt.y, 0, g.hwndMain, nullptr);
    DestroyMenu(m);
    if (cmd <= 0) return;
    for (auto& path : sel) {
        if (cmd == 8) g.colorTags.erase(path);
        else g.colorTags[path] = colors[cmd-1].c;
    }
    InvalidateRect(p.hwndList, nullptr, FALSE);
}

// =====================================================================
// SETTINGS DIALOG  (left-nav panel design)
// =====================================================================

// Additional control IDs for new settings layout
#define ID_SETT_NAV_BASE  3100   // 3100-3109: nav category buttons
#define ID_SETT_PAGE_AREA 3200   // page content area (custom paint)
#define ID_SETT_CAT_APPEARANCE 0
#define ID_SETT_CAT_LAYOUT     1
#define ID_SETT_CAT_BEHAVIOR   2
#define ID_SETT_CAT_BOOKMARKS  3

static void ShowSettingsDialog() {
    if (g.hwndSettings && IsWindow(g.hwndSettings)) {
        SetForegroundWindow(g.hwndSettings);
        return;
    }

    const int DW = 560, DH = 520;

    // --- Layout constants ---
    const int NAV_W   = 140;    // left nav panel width
    const int PAD     = 16;
    const int RX      = NAV_W + PAD;   // right content x
    const int RW      = DW - RX - PAD; // right content width
    const int EH      = 24;
    const int RH      = 36;
    const int BTN_H   = 30;
    const int BTN_W   = 90;
    const int FOOT_H  = 52;

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME, L"STATIC", L"OrbitFM \u2014 Settings",
        WS_POPUP|WS_CAPTION|WS_SYSMENU|WS_CLIPCHILDREN,
        0, 0, DW, DH, g.hwndMain, nullptr, g.hInst, nullptr);
    if (!dlg) return;
    g.hwndSettings = dlg;
    RECT rc; GetWindowRect(g.hwndMain, &rc);
    SetWindowPos(dlg, nullptr,
        rc.left + (rc.right  - rc.left - DW) / 2,
        rc.top  + (rc.bottom - rc.top  - DH) / 2,
        0, 0, SWP_NOSIZE|SWP_NOZORDER);

    // --- Helper lambdas ---
    int* pActiveCat = new int(ID_SETT_CAT_APPEARANCE);

    auto mkNavBtn = [&](int cat, const wchar_t* label) {
        int iy = 56 + cat * 42;
        HWND hw = CreateWindowExW(0, L"BUTTON", label,
            WS_CHILD|WS_VISIBLE|BS_OWNERDRAW,
            4, iy, NAV_W - 8, 36, dlg,
            (HMENU)(INT_PTR)(ID_SETT_NAV_BASE + cat), g.hInst, nullptr);
        SendMessageW(hw, WM_SETFONT, (WPARAM)g.hFont, TRUE);
    };
    auto mkLbl = [&](int id, const wchar_t* t, int x, int y, int w, int h, bool bold=false, bool show=true) {
        HWND hw = CreateWindowExW(0, L"STATIC", t, WS_CHILD|(show?WS_VISIBLE:0)|SS_LEFT,
            x, y, w, h, dlg, (HMENU)(INT_PTR)id, g.hInst, nullptr);
        SendMessageW(hw, WM_SETFONT, (WPARAM)(bold ? g.hFontBold : g.hFont), TRUE);
        return hw;
    };
    auto mkEdit = [&](int id, const std::wstring& v, int x, int y, int w, bool show=true) {
        HWND hw = CreateWindowExW(0, L"EDIT", v.c_str(),
            WS_CHILD|(show?WS_VISIBLE:0)|ES_AUTOHSCROLL,
            x, y, w, EH, dlg, (HMENU)(INT_PTR)id, g.hInst, nullptr);
        SendMessageW(hw, WM_SETFONT, (WPARAM)g.hFont, TRUE);
        return hw;
    };
    auto mkChk = [&](int id, const wchar_t* t, bool checked, int x, int y, int w, bool show=true) {
        HWND hw = CreateWindowExW(0, L"BUTTON", t,
            WS_CHILD|(show?WS_VISIBLE:0)|BS_AUTOCHECKBOX,
            x, y, w, 22, dlg, (HMENU)(INT_PTR)id, g.hInst, nullptr);
        SendMessageW(hw, WM_SETFONT, (WPARAM)g.hFont, TRUE);
        if (checked) SendMessageW(hw, BM_SETCHECK, BST_CHECKED, 0);
        return hw;
    };
    auto mkBtn = [&](int id, const wchar_t* t, int x, int y, int w, int h) {
        HWND hw = CreateWindowExW(0, L"BUTTON", t,
            WS_CHILD|WS_VISIBLE|BS_OWNERDRAW,
            x, y, w, h, dlg, (HMENU)(INT_PTR)id, g.hInst, nullptr);
        SendMessageW(hw, WM_SETFONT, (WPARAM)g.hFont, TRUE);
        return hw;
    };

    // --- Nav buttons ---
    mkNavBtn(ID_SETT_CAT_APPEARANCE, L"Appearance");
    mkNavBtn(ID_SETT_CAT_LAYOUT,     L"Layout");
    mkNavBtn(ID_SETT_CAT_BEHAVIOR,   L"Behavior");
    mkNavBtn(ID_SETT_CAT_BOOKMARKS,  L"Bookmarks");

    // Label ID allocation:
    //  3000 = "APPEARANCE" header,  3010 = "Font size:" label
    //  3001 = "LAYOUT" header,      3011-3013 = layout sub-labels
    //  3002 = "BEHAVIOR" header
    //  3003 = "BOOKMARKS" header,   3014 = hint line, 3015-3019 = slot labels

    // --- Page: APPEARANCE (cat 0) ---
    int y = 48;
    mkLbl(3000, L"APPEARANCE", RX, y, RW, 18, true); y += 28;
    mkLbl(3010, L"Font size (pt):", RX, y+5, 120, 18);
    mkEdit(ID_SET_FONT_SIZE, std::to_wstring(g.settingFontSize), RX+126, y, 60);
    y += RH;

    // --- Page: LAYOUT (cat 1) - initially hidden ---
    int pageY1 = 48;
    mkLbl(3001, L"LAYOUT", RX, pageY1, RW, 18, true, false); pageY1 += 28;
    mkLbl(3011, L"Sidebar width:", RX, pageY1+5, 120, 18, false, false);
    mkEdit(ID_SET_SIDEBAR_W, std::to_wstring(g.sidebarW), RX+126, pageY1, 72, false); pageY1 += RH;
    mkLbl(3012, L"Preview width:", RX, pageY1+5, 120, 18, false, false);
    mkEdit(ID_SET_PREVIEW_W, std::to_wstring(g.previewW), RX+126, pageY1, 72, false); pageY1 += RH;
    mkLbl(3013, L"Start path:", RX, pageY1+5, 120, 18, false, false);
    mkEdit(ID_SET_START_PATH, g.settingStartPath, RX+126, pageY1, RW-130, false); pageY1 += RH;

    // --- Page: BEHAVIOR (cat 2) ---
    int pageY2 = 48;
    mkLbl(3002, L"BEHAVIOR", RX, pageY2, RW, 18, true, false); pageY2 += 28;
    mkChk(ID_SET_SHOW_HIDDEN,  L"Show hidden files by default",  g.showHidden,        RX, pageY2, RW, false); pageY2 += 34;
    mkChk(ID_SET_AUTO_REFRESH, L"Auto-refresh on file changes",  g.settingAutoRefresh, RX, pageY2, RW, false); pageY2 += 34;

    // --- Page: BOOKMARKS (cat 3) ---
    int pageY3 = 48;
    mkLbl(3003, L"BOOKMARKS", RX, pageY3, RW, 18, true, false); pageY3 += 20;
    mkLbl(3014, L"Ctrl+Shift+1\u20135 sets  \u00b7  Ctrl+5\u20139 jumps", RX, pageY3, RW, 16, false, false);
    pageY3 += 24;
    for (int i = 0; i < 5; i++) {
        wchar_t t[16]; StringCchPrintfW(t, 16, L"Slot %d:", i+1);
        mkLbl(3015+i, t, RX, pageY3+5, 60, 18, false, false);
        mkEdit(2100+i, g.bookmarks[i], RX+66, pageY3, RW-70, false); pageY3 += RH;
    }

    // --- Footer buttons ---
    mkBtn(ID_SET_CANCEL, L"Cancel", DW - PAD - BTN_W,            DH - FOOT_H + 11, BTN_W, BTN_H);
    mkBtn(ID_SET_OK,     L"Save",   DW - PAD - BTN_W*2 - 8,      DH - FOOT_H + 11, BTN_W, BTN_H);

    // --- Window proc ---
    struct SD {
        static void ShowPage(HWND h, int cat, int navW) {
            // Map each control ID to its category
            // We store category in control's user data via SetPropW not available here,
            // so we hide/show based on position: controls in the right panel only
            // Use a simpler approach: enumerate and check x position
            // Controls with x >= navW are page controls; we manage them by tag
            // Since we can't tag, rebuild visibility by examining IDs:
            struct Range { int lo, hi, page; };
            static const Range ranges[] = {
                {3000, 3000, 0},                             // appearance header
                {3010, 3010, 0},                             // appearance sub-label
                {ID_SET_FONT_SIZE, ID_SET_FONT_SIZE, 0},     // appearance edit
                {3001, 3001, 1},                             // layout header
                {3011, 3013, 1},                             // layout sub-labels
                {ID_SET_SIDEBAR_W, ID_SET_START_PATH, 1},    // layout edits
                {3002, 3002, 2},                             // behavior header
                {ID_SET_SHOW_HIDDEN, ID_SET_AUTO_REFRESH, 2},// behavior checks
                {3003, 3003, 3},                             // bookmarks header
                {3014, 3019, 3},                             // bookmarks hint + slot labels
                {2100, 2104, 3},                             // bookmarks edits
            };
            HWND ch = GetWindow(h, GW_CHILD);
            while (ch) {
                int id = GetDlgCtrlID(ch);
                // skip nav buttons and footer buttons (they're always visible)
                bool isNav    = (id >= ID_SETT_NAV_BASE && id < ID_SETT_NAV_BASE + 10);
                bool isFooter = (id == ID_SET_OK || id == ID_SET_CANCEL);
                if (!isNav && !isFooter && id > 0) {
                    // Determine which page this control belongs to
                    int pg = -1;
                    for (auto& r : ranges) if (id >= r.lo && id <= r.hi) { pg = r.page; break; }
                    if (pg >= 0) ShowWindow(ch, (pg == cat) ? SW_SHOW : SW_HIDE);
                }
                ch = GetWindow(ch, GW_HWNDNEXT);
            }
            // Also hide/show the STATIC labels (which have id=0 — we can't distinguish them)
            // Accept: static labels without IDs are always visible; that's fine since
            // they use x >= navW as their position filter
        }

        static COLORREF DrawNavBtn(HDC dc, RECT r, const wchar_t* label,
                                   bool active, bool hover)
        {
            COLORREF bg = active ? COL_ACCENT : (hover ? COL_HOVER : COL_SIDEBAR);
            HBRUSH br = CreateSolidBrush(bg);
            FillRect(dc, &r, br); DeleteObject(br);
            // Left accent strip for active
            if (active) {
                HBRUSH ab = CreateSolidBrush(COL_ACCENT);
                RECT strip = {r.left, r.top, r.left+3, r.bottom};
                FillRect(dc, &strip, ab); DeleteObject(ab);
            }
            SetTextColor(dc, active ? RGB(255,255,255) : COL_TEXT);
            SetBkMode(dc, TRANSPARENT);
            RECT tr = {r.left+12, r.top, r.right-4, r.bottom};
            DrawTextW(dc, label, -1, &tr, DT_LEFT|DT_VCENTER|DT_SINGLELINE);
            return bg;
        }

        static LRESULT CALLBACK P(HWND h, UINT m, WPARAM wp, LPARAM lp) {
            int* pCat = (int*)GetWindowLongPtrW(h, GWLP_USERDATA);
            switch (m) {

            case WM_ERASEBKGND: {
                RECT r; GetClientRect(h, &r);
                HDC dc = (HDC)wp;
                // Main background
                HBRUSH br = CreateSolidBrush(COL_BG); FillRect(dc, &r, br); DeleteObject(br);
                // Left nav panel
                RECT nav = {0, 0, 140, r.bottom};
                HBRUSH nb = CreateSolidBrush(COL_SIDEBAR); FillRect(dc, &nav, nb); DeleteObject(nb);
                // Divider
                HPEN pen = CreatePen(PS_SOLID, 1, COL_BORDER);
                HPEN op = (HPEN)SelectObject(dc, pen);
                MoveToEx(dc, 140, 0, nullptr); LineTo(dc, 140, r.bottom);
                // Footer divider
                int footY = r.bottom - 52;
                MoveToEx(dc, 141, footY, nullptr); LineTo(dc, r.right, footY);
                SelectObject(dc, op); DeleteObject(pen);
                // Nav header "SETTINGS" label
                SetTextColor(dc, COL_ACCENT);
                SetBkMode(dc, TRANSPARENT);
                HFONT oldF = (HFONT)SelectObject(dc, (HFONT)SendMessageW(h, WM_GETFONT, 0, 0));
                RECT nh = {8, 16, 134, 44};
                DrawTextW(dc, L"SETTINGS", -1, &nh, DT_LEFT|DT_VCENTER|DT_SINGLELINE);
                SelectObject(dc, oldF);
                return 1;
            }

            case WM_PAINT: {
                PAINTSTRUCT ps; HDC dc = BeginPaint(h, &ps);
                RECT cr; GetClientRect(h, &cr);
                // Draw 1px border around visible EDIT children
                HBRUSH bBorder = CreateSolidBrush(COL_BORDER);
                HWND ch = GetWindow(h, GW_CHILD);
                while (ch) {
                    wchar_t cls[32] = {}; GetClassNameW(ch, cls, 32);
                    if (_wcsicmp(cls, L"EDIT") == 0 && IsWindowVisible(ch)) {
                        RECT er; GetWindowRect(ch, &er);
                        MapWindowPoints(nullptr, h, (POINT*)&er, 2);
                        InflateRect(&er, 1, 1);
                        FrameRect(dc, &er, bBorder);
                    }
                    ch = GetWindow(ch, GW_HWNDNEXT);
                }
                DeleteObject(bBorder);
                EndPaint(h, &ps);
                return 0;
            }

            case WM_CTLCOLORSTATIC: {
                HDC dc = (HDC)wp; HWND ctrl = (HWND)lp;
                RECT cr; GetClientRect(ctrl, &cr);
                POINT pt{0,0}; MapWindowPoints(ctrl, h, &pt, 1);
                // Controls in the nav panel get sidebar bg; right panel get main bg
                bool inNav = (pt.x < 140);
                COLORREF bg = inNav ? COL_SIDEBAR : COL_BG;
                int id = GetDlgCtrlID(ctrl);
                SetBkColor(dc, bg);
                if (id >= ID_SETT_SECT_BASE && id < ID_SETT_SECT_BASE + 20)
                    SetTextColor(dc, COL_ACCENT);
                else
                    SetTextColor(dc, inNav ? COL_DIM : COL_TEXT);
                static HBRUSH sBg = nullptr, sNav = nullptr;
                if (!sBg) sBg = CreateSolidBrush(COL_BG);
                if (!sNav) sNav = CreateSolidBrush(COL_SIDEBAR);
                return (LRESULT)(inNav ? sNav : sBg);
            }

            case WM_CTLCOLOREDIT: {
                HDC dc = (HDC)wp;
                SetTextColor(dc, COL_TEXT);
                SetBkColor(dc, COL_BG_ALT);
                static HBRUSH sEdit = nullptr;
                if (!sEdit) sEdit = CreateSolidBrush(COL_BG_ALT);
                return (LRESULT)sEdit;
            }

            case WM_CTLCOLORBTN: {
                HWND ctrl = (HWND)lp;
                POINT pt{0,0}; MapWindowPoints(ctrl, h, &pt, 1);
                bool inNav = (pt.x < 140);
                COLORREF bg = inNav ? COL_SIDEBAR : COL_BG;
                SetBkColor((HDC)wp, bg);
                static HBRUSH sBg = nullptr, sNav = nullptr;
                if (!sBg) sBg = CreateSolidBrush(COL_BG);
                if (!sNav) sNav = CreateSolidBrush(COL_SIDEBAR);
                return (LRESULT)(inNav ? sNav : sBg);
            }

            case WM_DRAWITEM: {
                auto* di = (DRAWITEMSTRUCT*)lp;
                int id = (int)di->CtlID;
                bool pressed = (di->itemState & ODS_SELECTED) != 0;
                wchar_t txt[64] = {}; GetWindowTextW(di->hwndItem, txt, 63);

                if (id >= ID_SETT_NAV_BASE && id < ID_SETT_NAV_BASE + 10) {
                    // Nav button
                    int cat = id - ID_SETT_NAV_BASE;
                    bool active = (pCat && *pCat == cat);
                    DrawNavBtn(di->hDC, di->rcItem, txt, active, pressed && !active);
                    HFONT hf = (HFONT)SendMessageW(di->hwndItem, WM_GETFONT, 0, 0);
                    HFONT ohf = (HFONT)SelectObject(di->hDC, hf);
                    SetTextColor(di->hDC, active ? RGB(255,255,255) : COL_TEXT);
                    SetBkMode(di->hDC, TRANSPARENT);
                    RECT tr = {di->rcItem.left + 12, di->rcItem.top,
                               di->rcItem.right - 4, di->rcItem.bottom};
                    DrawTextW(di->hDC, txt, -1, &tr, DT_LEFT|DT_VCENTER|DT_SINGLELINE);
                    SelectObject(di->hDC, ohf);
                    return TRUE;
                }

                // Footer buttons (Save / Cancel)
                bool isSave = (id == ID_SET_OK);
                COLORREF bg = isSave ? (pressed ? COL_SEL : COL_ACCENT)
                                     : (pressed ? COL_HOVER : COL_BG_ALT);
                COLORREF border = isSave ? COL_ACCENT : COL_BORDER;
                COLORREF fg = isSave ? RGB(255,255,255) : COL_TEXT;
                HBRUSH br = CreateSolidBrush(bg); FillRect(di->hDC, &di->rcItem, br); DeleteObject(br);
                HPEN pen = CreatePen(PS_SOLID, 1, border);
                HPEN op  = (HPEN)SelectObject(di->hDC, pen);
                HBRUSH nb = (HBRUSH)GetStockObject(NULL_BRUSH);
                HBRUSH ob = (HBRUSH)SelectObject(di->hDC, nb);
                Rectangle(di->hDC, di->rcItem.left, di->rcItem.top,
                                   di->rcItem.right, di->rcItem.bottom);
                SelectObject(di->hDC, op); DeleteObject(pen);
                SelectObject(di->hDC, ob);
                SetTextColor(di->hDC, fg);
                SetBkMode(di->hDC, TRANSPARENT);
                HFONT hf  = (HFONT)SendMessageW(di->hwndItem, WM_GETFONT, 0, 0);
                HFONT ohf = (HFONT)SelectObject(di->hDC, hf);
                DrawTextW(di->hDC, txt, -1, &di->rcItem, DT_CENTER|DT_VCENTER|DT_SINGLELINE);
                SelectObject(di->hDC, ohf);
                return TRUE;
            }

            case WM_COMMAND: {
                int id = LOWORD(wp);
                // Nav category switch
                if (id >= ID_SETT_NAV_BASE && id < ID_SETT_NAV_BASE + 10) {
                    int newCat = id - ID_SETT_NAV_BASE;
                    if (pCat) *pCat = newCat;
                    ShowPage(h, newCat, 140);
                    InvalidateRect(h, nullptr, TRUE);
                    return 0;
                }
                if (id == ID_SET_OK) {
                    wchar_t buf[MAX_PATH] = {};
                    GetDlgItemTextW(h, ID_SET_FONT_SIZE,  buf, 511);
                    int fs = _wtoi(buf); if (fs>=8 && fs<=24) g.settingFontSize = fs;
                    GetDlgItemTextW(h, ID_SET_SIDEBAR_W,  buf, 511);
                    int sw = _wtoi(buf); if (sw>=60 && sw<=600) g.sidebarW = sw;
                    GetDlgItemTextW(h, ID_SET_PREVIEW_W,  buf, 511);
                    int pw = _wtoi(buf); if (pw>=60 && pw<=800) g.previewW = pw;
                    GetDlgItemTextW(h, ID_SET_START_PATH, buf, MAX_PATH-1);
                    g.settingStartPath = buf;
                    g.showHidden        = IsDlgButtonChecked(h, ID_SET_SHOW_HIDDEN)  == BST_CHECKED;
                    g.settingAutoRefresh= IsDlgButtonChecked(h, ID_SET_AUTO_REFRESH) == BST_CHECKED;
                    for (int i = 0; i < 5; i++) {
                        GetDlgItemTextW(h, 2100+i, buf, MAX_PATH-1);
                        g.bookmarks[i] = buf;
                    }
                    SaveSettings();
                    LayoutMain(g.hwndMain);
                    for (auto& pn : g.panes) RefreshPane(pn);
                    DestroyWindow(h);
                } else if (id == ID_SET_CANCEL) {
                    DestroyWindow(h);
                }
                return 0;
            }

            case WM_CLOSE:
                DestroyWindow(h);
                return 0;
            case WM_DESTROY: {
                int* pc = (int*)GetWindowLongPtrW(h, GWLP_USERDATA);
                delete pc;
                g.hwndSettings = nullptr;
                for (auto& pn : g.panes)
                    if (pn.hwnd) InvalidateRect(pn.hwnd, nullptr, FALSE);
                return 0;
            }
            }
            return DefWindowProcW(h, m, wp, lp);
        }
    };
    SetWindowLongPtrW(dlg, GWLP_WNDPROC, (LONG_PTR)SD::P);
    SetWindowLongPtrW(dlg, GWLP_USERDATA, (LONG_PTR)pActiveCat);
    SendMessageW(dlg, WM_SETFONT, (WPARAM)g.hFontBold, FALSE);

    // Show Appearance page by default
    SD::ShowPage(dlg, ID_SETT_CAT_APPEARANCE, NAV_W);

    ShowWindow(dlg, SW_SHOW);
    for (auto& pn : g.panes)
        if (pn.hwnd) InvalidateRect(pn.hwnd, nullptr, FALSE);
}

// =====================================================================
// MAIN WINDOW LAYOUT + ACTIVE PANE
// =====================================================================

static void LayoutMain(HWND hwnd) {
    RECT cr; GetClientRect(hwnd, &cr);
    int W = cr.right;
    int H = cr.bottom;
    int statH = g.hwndStatus ? STATUS_H : 0;
    int contentH = H - statH;
    int sb = g.sidebarW;
    int pv = g.previewW;
    if (sb < 60) sb = 60;
    if (pv < 60) pv = 60;
    int paneTotal = W - sb - pv;
    if (paneTotal < 200) {
        pv = std::max(60, W - sb - 200);
        paneTotal = W - sb - pv;
        if (paneTotal < 100) { sb = 100; paneTotal = W - sb - pv; }
    }
    int paneX = sb;
    int p1w, p2w;
    if (g.showPane2) {
        p1w = paneTotal / 2;
        p2w = paneTotal - p1w;
    } else {
        p1w = paneTotal;
        p2w = 0;
    }
    MoveWindow(g.hwndSidebar, 0, 0, sb, contentH, TRUE);
    MoveWindow(g.panes[0].hwnd, paneX, 0, p1w, contentH, TRUE);
    if (g.panes[1].hwnd) {
        if (g.showPane2)
            MoveWindow(g.panes[1].hwnd, paneX + p1w, 0, p2w, contentH, TRUE);
        else
            SetWindowPos(g.panes[1].hwnd, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOZORDER | SWP_HIDEWINDOW);
    }
    MoveWindow(g.hwndPreview, paneX + p1w + p2w, 0, pv, contentH, TRUE);
    if (g.hwndStatus) MoveWindow(g.hwndStatus, 0, contentH, W, statH, TRUE);
}

static void SetActivePane(int idx) {
    if (idx < 0 || idx > 1) return;
    if (idx == 1 && !g.showPane2) idx = 0;
    if (g.activePane == idx) return;
    g.activePane = idx;
    // Start pulse on newly active pane, stop on the other
    for (int pi = 0; pi < 2; pi++) {
        if (g.panes[pi].hwnd) {
            if (pi == idx) SetTimer(g.panes[pi].hwnd, TIMER_PULSE + pi, 40, nullptr);
            else           KillTimer(g.panes[pi].hwnd, TIMER_PULSE + pi);
        }
    }
    InvalidateRect(g.panes[0].hwnd, nullptr, FALSE);
    InvalidateRect(g.panes[1].hwnd, nullptr, FALSE);
    int focIdx = ListView_GetNextItem(g.panes[idx].hwndList, -1, LVNI_FOCUSED);
    if (focIdx >= 0 && !g.panes[idx].tabs.empty() &&
        focIdx < (int)g.panes[idx].tabs[g.panes[idx].activeTab].entries.size()) {
        auto& e = g.panes[idx].tabs[g.panes[idx].activeTab].entries[focIdx];
        if (!e.isDir())
            UpdatePreview(CombinePath(g.panes[idx].tabs[g.panes[idx].activeTab].path, e.name));
        else
            UpdatePreview(L"");
    } else {
        UpdatePreview(L"");
    }
}

// =====================================================================
// MAIN WINDOW PROC
// =====================================================================

static LRESULT CALLBACK MainProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
        case WM_CREATE: {
            // Check for tear-off path embedded in window title: "OrbitFM|C:\path"
            std::wstring initPath;
            {
                wchar_t titleBuf[MAX_PATH + 32]{};
                GetWindowTextW(hwnd, titleBuf, MAX_PATH + 31);
                std::wstring title = titleBuf;
                size_t pipe = title.find(L'|');
                if (pipe != std::wstring::npos) {
                    initPath = title.substr(pipe + 1);
                    SetWindowTextW(hwnd, L"OrbitFM");
                }
            }

            // For tear-off windows, we only have one logical instance of g.
            // We detect this is a secondary window if g.hwndMain is already set.
            bool isSecondary = (g.hwndMain != nullptr);
            if (!isSecondary) g.hwndMain = hwnd;

            if (!isSecondary) {
                g.hwndSidebar = CreateWindowExW(0, CLASS_SIDEBAR, L"",
                    WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g.hInst, nullptr);
                for (int i = 0; i < 2; i++) {
                    g.panes[i].paneIndex = i;
                    g.panes[i].tabs.clear();
                    PaneTab t;
                    t.path = (i == 0) ? KnownFolderById(KFID_Desktop) : KnownFolderById(KFID_Documents);
                    if (t.path.empty() || !DirExists(t.path)) t.path = KnownFolderById(KFID_Profile);
                    if (t.path.empty() || !DirExists(t.path)) t.path = L"C:\\";
                    g.panes[i].tabs.push_back(t);
                    CreateWindowExW(0, CLASS_PANE, L"",
                        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g.hInst, &g.panes[i]);
                }
                g.hwndPreview = CreateWindowExW(0, CLASS_PREVIEW, L"",
                    WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g.hInst, nullptr);
                g.hwndStatus = CreateWindowExW(0, CLASS_STATUS, L"",
                    WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, nullptr, g.hInst, nullptr);
            } else {
                // Tear-off secondary window: create a minimal independent shell
                // that opens Explorer at the dropped tab's path
                if (!initPath.empty() && DirExists(initPath))
                    ShellExecuteW(nullptr, L"explore", initPath.c_str(), nullptr, nullptr, SW_SHOW);
                DestroyWindow(hwnd);
                return 0;
            }
            return 0;
        }
        case WM_SHOWWINDOW: {
            if (w) {
                static bool firstShow = true;
                if (firstShow) {
                    firstShow = false;
                    // Try to restore session
                    LoadSession();
                    // If session loaded tabs for pane 0, navigate to them
                    for (int pi = 0; pi < 2; pi++) {
                        if (!g.panes[pi].tabs.empty()) {
                            NavigateTabInternal(g.panes[pi], g.panes[pi].activeTab,
                                g.panes[pi].tabs[g.panes[pi].activeTab].path, false);
                        }
                    }
                    // If start path is configured, use it for pane 0
                    if (!g.settingStartPath.empty() && DirExists(g.settingStartPath) && g.panes[0].tabs.empty())
                        NavigateTab(g.panes[0], 0, g.settingStartPath);
                    RefreshPane(g.panes[0]);
                    if (g.showPane2) RefreshPane(g.panes[1]);
                    SetActivePane(0);
                    SetFocus(g.panes[0].hwndList);
                    // Start auto-refresh watchers
                    if (g.settingAutoRefresh) {
                        StartWatchThread(0);
                        if (g.showPane2) StartWatchThread(1);
                    }
                    // Start git status scan
                    if (!g.panes[0].tabs.empty())
                        StartGitStatus(g.panes[0].tabs[g.panes[0].activeTab].path);
                }
            }
            return 0;
        }
        case WM_SIZE:
            LayoutMain(hwnd);
            return 0;
        case WM_MOUSEMOVE: {
            int mx = GET_X_LPARAM(l);
            if (g.splitDrag >= 0) {
                int delta = mx - g.splitDragStart.x;
                if (g.splitDrag == 0) {
                    g.sidebarW = std::max(80, g.splitDragOrig + delta);
                } else if (g.splitDrag == 1) {
                    g.previewW = std::max(80, g.splitDragOrig - delta);
                }
                LayoutMain(hwnd);
                return 0;
            }
            if (g.dragging) {
                if (g.dragArmed) {
                    POINT cp; GetCursorPos(&cp);
                    if (abs(cp.x - g.dragStart.x) < GetSystemMetrics(SM_CXDRAG) &&
                        abs(cp.y - g.dragStart.y) < GetSystemMetrics(SM_CYDRAG))
                        return 0;
                    g.dragArmed = false;
                }
                bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                SetCursor(LoadCursor(nullptr, ctrl ? IDC_CROSS : IDC_HAND));
                return 0;
            }
            // Splitter hit-test for cursor
            {
                RECT cr; GetClientRect(hwnd, &cr);
                int sb = g.sidebarW, pv = g.previewW;
                int paneTotal = cr.right - sb - pv;
                int p1w = paneTotal / 2;
                bool onSplit = (mx >= sb - SPLIT_W && mx <= sb + SPLIT_W) ||
                               (mx >= cr.right - pv - SPLIT_W && mx <= cr.right - pv + SPLIT_W);
                if (onSplit) SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
            }
            return 0;
        }
        case WM_LBUTTONDOWN: {
            int mx = GET_X_LPARAM(l);
            RECT cr; GetClientRect(hwnd, &cr);
            int sb = g.sidebarW, pv = g.previewW;
            if (mx >= sb - SPLIT_W && mx <= sb + SPLIT_W) {
                g.splitDrag = 0;
                g.splitDragStart = { mx, 0 };
                g.splitDragOrig = g.sidebarW;
                SetCapture(hwnd);
                return 0;
            }
            if (mx >= cr.right - pv - SPLIT_W && mx <= cr.right - pv + SPLIT_W) {
                g.splitDrag = 1;
                g.splitDragStart = { mx, 0 };
                g.splitDragOrig = g.previewW;
                SetCapture(hwnd);
                return 0;
            }
            break;
        }
        case WM_LBUTTONUP: {
            if (g.splitDrag >= 0) {
                g.splitDrag = -1;
                ReleaseCapture();
                return 0;
            }
            if (g.dragging) {
                g.dragging = false;
                ReleaseCapture();
                POINT cp; GetCursorPos(&cp);
                HWND hw = WindowFromPoint(cp);
                int destPane = -1;
                while (hw && hw != hwnd) {
                    if (hw == g.panes[0].hwnd || hw == g.panes[0].hwndList) { destPane = 0; break; }
                    if (hw == g.panes[1].hwnd || hw == g.panes[1].hwndList) { destPane = 1; break; }
                    hw = GetParent(hw);
                }
                if (destPane >= 0 && destPane != g.dragSourcePane && !g.dragFiles.empty()) {
                    bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                    std::wstring dest = g.panes[destPane].tabs[g.panes[destPane].activeTab].path;
                    if (ctrl) CopyFiles(g.hwndMain, g.dragFiles, dest);
                    else      MoveFiles(g.hwndMain, g.dragFiles, dest);
                    RefreshPane(g.panes[0]);
                    RefreshPane(g.panes[1]);
                }
                g.dragFiles.clear();
                g.dragSourcePane = -1;
            }
            return 0;
        }
        case WM_CAPTURECHANGED: {
            g.splitDrag = -1;
            if (g.dragging) {
                g.dragging = false;
                g.dragFiles.clear();
                g.dragSourcePane = -1;
            }
            return 0;
        }
        case WM_SETFOCUS: {
            if (g.activePane >= 0 && g.activePane < 2)
                SetFocus(g.panes[g.activePane].hwndList);
            return 0;
        }
        case WM_GETMINMAXINFO: {
            MINMAXINFO* mm = (MINMAXINFO*)l;
            mm->ptMinTrackSize.x = 800;
            mm->ptMinTrackSize.y = 500;
            return 0;
        }
        case WM_DUPE_READY:
            ShowDupeResults();
            return 0;
        case WM_DESTROY:
            StopWatchThreads();
            g.searchBgStop.store(true);
            SaveSession();
            SaveSettings();
            PostQuitMessage(0);
            return 0;
        case WM_ERASEBKGND: return 1;
    }
    return DefWindowProcW(hwnd, msg, w, l);
}

// =====================================================================
// REGISTRATION
// =====================================================================

static ATOM RegisterClass(const wchar_t* name, WNDPROC proc, HCURSOR cursor) {
    WNDCLASSEXW wc{ sizeof(wc) };
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = proc;
    wc.hInstance = g.hInst;
    wc.hCursor = cursor ? cursor : LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = name;
    return RegisterClassExW(&wc);
}

// =====================================================================
// wWinMain
// =====================================================================

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nCmdShow) {
    g.hInst = hInst;
    OleInitialize(nullptr);
    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    LOGFONTW lf{};
    HDC dc = GetDC(nullptr);
    lf.lfHeight = -MulDiv(10, GetDeviceCaps(dc, LOGPIXELSY), 72);
    lf.lfWeight = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfQuality = CLEARTYPE_QUALITY;
    StringCchCopyW(lf.lfFaceName, LF_FACESIZE, L"Segoe UI");
    g.hFont = CreateFontIndirectW(&lf);
    lf.lfWeight = FW_BOLD;
    g.hFontBold = CreateFontIndirectW(&lf);

    // Monospace font for text preview
    {
        LOGFONTW mf{};
        mf.lfHeight = -MulDiv(9, GetDeviceCaps(dc, LOGPIXELSY), 72);
        mf.lfWeight = FW_NORMAL;
        mf.lfCharSet = DEFAULT_CHARSET;
        mf.lfQuality = CLEARTYPE_QUALITY;
        mf.lfPitchAndFamily = FIXED_PITCH | FF_MODERN;
        StringCchCopyW(mf.lfFaceName, LF_FACESIZE, L"Consolas");
        g.hFontMono = CreateFontIndirectW(&mf);
    }
    ReleaseDC(nullptr, dc);

    // Enumerate local drives for sidebar THIS PC section
    {
        wchar_t driveBuf[256] = {};
        DWORD len = GetLogicalDriveStringsW(255, driveBuf);
        wchar_t* p = driveBuf;
        while (p < driveBuf + len && *p) {
            UINT t = GetDriveTypeW(p);
            if (t == DRIVE_FIXED || t == DRIVE_REMOVABLE || t == DRIVE_REMOTE || t == DRIVE_CDROM)
                g.drives.push_back(p);
            p += wcslen(p) + 1;
        }
    }

    // Detect OneDrive path from environment
    {
        wchar_t odBuf[MAX_PATH] = {};
        if (GetEnvironmentVariableW(L"OneDrive", odBuf, MAX_PATH) > 0)
            g.oneDrivePath = odBuf;
        else if (GetEnvironmentVariableW(L"OneDriveConsumer", odBuf, MAX_PATH) > 0)
            g.oneDrivePath = odBuf;
    }

    LoadTheme();
    RebuildBrushes();

    {
        SHFILEINFOW sfi{};
        HIMAGELIST il = (HIMAGELIST)SHGetFileInfoW(L"C:\\", 0, &sfi, sizeof(sfi),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
        g.sysImageList = il;
    }

    g.sidebarW = SIDEBAR_W;
    g.previewW = PREVIEW_W;
    LoadSettings(); // Load before LoadPins since it may update sidebarW/previewW
    LoadPins();
    LoadRecent();
    LoadBookmarks();
    StartThumbThread();
    StartHashThread();

    RegisterClass(CLASS_MAIN, MainProc, nullptr);
    RegisterClass(CLASS_SIDEBAR, SidebarProc, nullptr);
    RegisterClass(CLASS_PANE, PaneProc, nullptr);
    RegisterClass(CLASS_PREVIEW, PreviewProc, nullptr);
    RegisterClass(CLASS_STATUS, StatusProc, nullptr);

    // Register new overlay window classes (simple DefWindowProc-based; procs set at creation)
    {
        WNDCLASSEXW wc{ sizeof(wc) };
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = DefWindowProcW;
        wc.hInstance = g.hInst;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW+1);
        wc.lpszClassName = CLASS_PALETTE;
        RegisterClassExW(&wc);
        wc.lpszClassName = CLASS_CMDPAL;
        RegisterClassExW(&wc);
        wc.lpszClassName = CLASS_QUICKOPEN;
        RegisterClassExW(&wc);
    }

    HWND hwnd = CreateWindowExW(0, CLASS_MAIN, L"OrbitFM",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 1280, 780,
        nullptr, nullptr, hInst, nullptr);

    if (!hwnd) return 1;

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        HWND focus = GetFocus();
        wchar_t cls[64] = { 0 };
        if (focus) GetClassNameW(focus, cls, 64);
        bool isEdit = (_wcsicmp(cls, L"Edit") == 0);
        if (!isEdit) {
            if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN) {
                if (msg.wParam == VK_F5 || msg.wParam == VK_F2 || msg.wParam == VK_F1) {
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                    continue;
                }
            }
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    StopWatchThreads();
    StopThumbThread();
    StopHashThread();
    g.sizeCancel.store(true);
    if (g.sizeThread.joinable()) g.sizeThread.join();
    g.searchBgStop.store(true);
    if (g.searchBgThread.joinable()) g.searchBgThread.join();
    g.dupeRunning.store(false);
    // dupeThread is detached — watchStop flag will let it wind down
    g_qoStop.store(true);
    if (g.hFont) DeleteObject(g.hFont);
    if (g.hFontBold) DeleteObject(g.hFontBold);
    if (g.hFontMono) DeleteObject(g.hFontMono);
    if (g.brBg) DeleteObject(g.brBg);
    if (g.brBgAlt) DeleteObject(g.brBgAlt);
    if (g.brSel) DeleteObject(g.brSel);
    if (g.brHeader) DeleteObject(g.brHeader);
    OleUninitialize();
    return (int)msg.wParam;
}
