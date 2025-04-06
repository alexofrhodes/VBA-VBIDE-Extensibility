@echo off
setlocal enabledelayedexpansion

:: Step 1: Locate AutoHotkey installation directory
set "AhkDir="
for /d %%D in ("%ProgramFiles%\AutoHotkey*") do set "AhkDir=%%D"

if not defined AhkDir (
    echo AutoHotkey installation not found. Please install it first.
    pause
    exit /b
)

echo AutoHotkey found at: %AhkDir%
REM pause

:: Step 2: Locate Ahk2Exe (v1 and v2 compilers)
set "Ahk2Exe=%AhkDir%\Compiler\Ahk2Exe.exe"
if not exist "%Ahk2Exe%" (
    echo Ahk2Exe compiler not found in "%AhkDir%\Compiler\".
    pause
    exit /b
)

echo Ahk2Exe compiler found at: %Ahk2Exe%
REM pause

:: Step 3: Locate AutoHotkey v1 and v2 compilers
set "Compiler-v1="
set "Compiler-v2="

:: Locate AutoHotkeyU64.exe in any subfolder starting with "v1"
for /d %%D in ("%AhkDir%\v1*") do (
    if exist "%%D\AutoHotkeyU64.exe" (
        set "Compiler-v1=%%D\AutoHotkeyU64.exe"
        goto :foundV1
    )
)

:foundV1

:: Locate AutoHotkey64.exe in the v2 folder
for /d %%D in ("%AhkDir%\v2*") do (
    if exist "%%D\AutoHotkey64.exe" (
        set "Compiler-v2=%%D\AutoHotkey64.exe"
        goto :foundV2
    )
)

:foundV2

:: Ensure both compilers are found
if not defined Compiler-v1 (
    echo Error: AutoHotkeyU64.exe not found in v1 folder.
    pause
    exit /b
)

if not defined Compiler-v2 (
    echo Error: AutoHotkey64.exe not found in v2 folder.
    pause
    exit /b
)

REM echo AutoHotkey v1 compiler found at: %Compiler-v1%
REM echo AutoHotkey v2 compiler found at: %Compiler-v2%
REM pause

:: Step 4: Process the AHK file selection
set "thisBatDir=%~dp0"
set "foundMultiple=false"
set "fileList="
set "index=0"

:: Loop through the .ahk files in the current directory
for %%F in ("%thisBatDir%*.ahk") do (
    set /a index+=1
    set "fileList=!fileList!%%F (File !index!)"
    set "file[!index!]=%%F"
    set "ico[!index!]=%%~nF.ico"
    set "foundMultiple=true"
)

:: Step 5: Ask for compiler choice before showing file list
:askCompilerChoice
set /p compilerChoice="Select compiler (1 for v1 - AutoHotkeyU64.exe, 2 for v2 - AutoHotkey64.exe): "

if "%compilerChoice%"=="1" (
    set "Compiler=%Compiler-v1%"
) else if "%compilerChoice%"=="2" (
    set "Compiler=%Compiler-v2%"
) else (
    echo Invalid choice. Please enter 1 for v1 or 2 for v2.
    goto askCompilerChoice
)

:: Step 6: If only one file found, compile it immediately
if "%index%"=="1" (
    set "ahkFile=!file[1]!"
    set "icoFile=!ico[1]!"
    echo Only one AHK file found: !ahkFile!. Compiling now...
    if exist "!icoFile!" (
        "%Ahk2Exe%" /in "!ahkFile!" /base "%Compiler%" /icon "!icoFile!"
    ) else (
        echo Icon file not found for !ahkFile!, compiling without icon...
        "%Ahk2Exe%" /in "!ahkFile!" /base "%Compiler%"
    )
    echo.
    pause
    exit /b
)

:: Step 7: If multiple files found, prompt the user to choose
if "%foundMultiple%"=="true" (
    echo Multiple AHK files found. Please specify which files to compile:
    echo.
    for /L %%i in (1,1,!index!) do (
        echo %%i: !file[%%i]!
    )

    echo.
    echo You can enter a space-separated list of numbers (e.g., 1 2) or type "a" to compile all files.
    :askFileChoice
    set /p fileChoice="Enter your choice: "

    :: If the user selects 'all', select all files
    if /i "%fileChoice%"=="a" (
        set fileChoice=1-%index%
    )

    :: Validate fileChoice input
    for %%i in (%fileChoice%) do (
        if "%%i" lss "1" (
            echo Invalid choice! Please enter numbers from 1 to %index%.
            goto askFileChoice
        )
    )

    :: Step 8: Compile the selected AHK files using the chosen compiler
    for %%i in (%fileChoice%) do (
        set "ahkFile=!file[%%i]!"
        set "icoFile=!ico[%%i]!"
        echo Compiling: !ahkFile!

        :: Check if the ICO file exists
        if exist "!icoFile!" (
            "%Ahk2Exe%" /in "!ahkFile!" /base "%Compiler%" /icon "!icoFile!"
        ) else (
            echo Icon file not found for !ahkFile!, compiling without icon...
            "%Ahk2Exe%" /in "!ahkFile!" /base "%Compiler%"
        )
    )
)

echo.
pause
