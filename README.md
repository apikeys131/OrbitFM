# OrbitFM

A fast, keyboard-driven file manager for Windows. Get where you're going without the bloat.

## Features

- Dual-pane layout — work across two directories simultaneously
- Fast fuzzy search — find files instantly as you type
- Batch rename and move operations
- Full keyboard navigation — every action reachable without a mouse
- Tabs and bookmarks for frequently used locations
- File preview support
- Copy, move, delete with progress indicators
- Context menu integration with Windows shell

## Download

Grab the latest installer from [Releases](https://github.com/apikeys131/OrbitFM/releases).

Run `OrbitFM-Setup-v1.0.0.exe` and follow the installer. Windows may show a SmartScreen warning — click **More info → Run anyway**.

## Requirements

- Windows 10 or 11 (64-bit)

## Building

```bash
# 1. Clone the repository
git clone https://github.com/apikeys131/OrbitFM.git
cd OrbitFM

# 2. Configure
cmake -B build -G "Visual Studio 17 2022" -A x64 -DCMAKE_BUILD_TYPE=Release

# 3. Build
cmake --build build --config Release --parallel

# 4. Run
.\build\Release\filemanager.exe
```

## Architecture

```
src/
├── main.cpp              Entry point
├── ui/
│   ├── MainWindow.h/.cpp Main window and layout
│   ├── Pane.h/.cpp       Individual file pane (left / right)
│   └── Toolbar.h/.cpp    Top toolbar and tab bar
├── fs/
│   ├── FileSystem.h/.cpp Directory listing, file ops (copy/move/delete)
│   └── Watcher.h/.cpp    File system change watcher
└── utils/
    ├── FuzzySearch.h     Fuzzy match algorithm
    └── Bookmarks.h/.cpp  Persistent bookmark storage
```

## Source

Source is available for review. See the repository for build instructions.

## License

Source-available. Free to use. Do not redistribute modified binaries without permission.
