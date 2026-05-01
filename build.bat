@echo off
setlocal EnableExtensions EnableDelayedExpansion
pushd "%~dp0"

set "OUT=filemanager.exe"
set "SRC=filemanager.cpp"

where g++ >nul 2>&1
if not errorlevel 1 (
    echo [build] Using g++
    g++ -std=c++17 -O2 -municode -mwindows -DUNICODE -D_UNICODE ^
        "%SRC%" -o "%OUT%" ^
        -lcomctl32 -lshell32 -lshlwapi -lole32 -loleaut32 -luuid ^
        -lpropsys -luxtheme -lgdi32 -luser32
    if errorlevel 1 (
        echo [build] g++ build failed.
        popd
        exit /b 1
    )
    echo [build] OK: %OUT%
    popd
    exit /b 0
)

where cl >nul 2>&1
if not errorlevel 1 (
    echo [build] Using MSVC cl
    cl /nologo /EHsc /O2 /std:c++17 /W3 /MT /DUNICODE /D_UNICODE "%SRC%" ^
        /link /SUBSYSTEM:WINDOWS /OUT:"%OUT%" ^
        comctl32.lib shell32.lib shlwapi.lib ole32.lib oleaut32.lib uuid.lib ^
        propsys.lib uxtheme.lib gdi32.lib user32.lib
    if errorlevel 1 (
        echo [build] MSVC build failed.
        popd
        exit /b 1
    )
    echo [build] OK: %OUT%
    popd
    exit /b 0
)

echo [build] No compiler found.
echo Install MinGW-w64 (g++) and add it to PATH, or run this from
echo the "x64 Native Tools Command Prompt" so cl.exe is available.
popd
exit /b 1
