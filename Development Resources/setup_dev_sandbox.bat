@echo off
setlocal EnableDelayedExpansion
title Development Environment Sandbox Launcher

:: ============================================================================
:: Launcher for Development Environment Windows Sandbox
:: Generates .wsb file with correct absolute paths and launches it
:: ============================================================================

echo ============================================================================
echo  Development Environment Sandbox Launcher
echo ============================================================================
echo.

:: Get the directory where this script is located
set SCRIPT_DIR=%~dp0
:: Remove trailing backslash
set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%

set SHARED_DIR=%SCRIPT_DIR%\shared
set OUTPUT_DIR=%SCRIPT_DIR%\VS Sandbox Output
set SPECFILE_DIR=%SHARED_DIR%\specfile_lv
set VS_LAYOUT_DIR=%SHARED_DIR%\VSLayout
set VS_LAYOUT=%VS_LAYOUT_DIR%\vs_community.exe
set VS_BOOTSTRAPPER=%SHARED_DIR%\vs_community.exe
set WSB_FILE=%SCRIPT_DIR%\install_dev_environment.wsb

:: Verify shared directory exists
if not exist "%SHARED_DIR%" (
    echo ERROR: Shared directory not found:
    echo   %SHARED_DIR%
    echo.
    echo Please run this script from the correct location.
    echo.
    pause
    exit /b 1
)

echo Shared folder: %SHARED_DIR%

:: Create output directory if it doesn't exist
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
echo Output folder: %OUTPUT_DIR%

:: ============================================================================
:: Check for Visual Studio offline layout
:: ============================================================================
if exist "%VS_LAYOUT%" goto :vs_layout_found

echo.
echo No Visual Studio offline layout found.
echo   Expected location: %VS_LAYOUT%
echo.
echo Options:
echo   [1] Continue without offline layout (will download ~2GB in sandbox each time^)
echo   [2] Create offline layout now (downloads ~8GB once, reusable^)
echo   [3] Exit
echo.
choice /c 123 /n /m "Select option (1, 2, or 3): "
if errorlevel 3 exit /b 0
if errorlevel 2 goto :create_vs_layout
goto :vs_layout_done

:create_vs_layout
echo.
echo Downloading Visual Studio bootstrapper...
curl -L -o "%VS_BOOTSTRAPPER%" --progress-bar "https://aka.ms/vs/stable/vs_community.exe"
if not exist "%VS_BOOTSTRAPPER%" (
    echo ERROR: Failed to download Visual Studio bootstrapper!
    pause
    exit /b 1
)

echo.
echo Creating offline layout at: %VS_LAYOUT_DIR%
echo This will download approximately 8GB. Please wait...
echo.
mkdir "%VS_LAYOUT_DIR%" 2>nul
:: Note: --remove is not supported by the bootstrapper, only by the layout installer
:: "%VS_BOOTSTRAPPER%" --layout "%VS_LAYOUT_DIR%" --add Microsoft.VisualStudio.Workload.NativeDesktop --add Microsoft.VisualStudio.Component.VC.Tools.x86.x64 --add Microsoft.VisualStudio.Component.Windows11SDK.22621 --add Microsoft.VisualStudio.Component.VC.CMake.Project --remove Component.GitHub.Copilot --remove Component.GitHub.Copilot.Chat --includeRecommended --lang en-US --wait
"%VS_BOOTSTRAPPER%" --layout "%VS_LAYOUT_DIR%" --add Microsoft.VisualStudio.Workload.NativeDesktop --add Microsoft.VisualStudio.Component.VC.Tools.x86.x64 --add Microsoft.VisualStudio.Component.Windows11SDK.22621 --add Microsoft.VisualStudio.Component.VC.CMake.Project --add Microsoft.VisualStudio.Component.Git --includeRecommended --lang en-US --wait

if exist "%VS_LAYOUT%" (
    echo.
    echo Offline layout created successfully!
) else (
    echo.
    echo WARNING: Layout creation may have failed. Check %VS_LAYOUT_DIR%
    pause
)
goto :vs_layout_done

:vs_layout_found
:vs_layout_done

echo.
echo Visual Studio layout:
if exist "%VS_LAYOUT%" (
    echo   Using offline layout: %VS_LAYOUT%
) else (
    echo   No offline layout (will download in sandbox^)
)

:: ============================================================================
:: Check for LabVIEW ISO and specfile
:: ============================================================================
set LABVIEW_ISO=%SHARED_DIR%\2015myRIO Software Bundle 1.iso

:check_labview_iso
if not exist "%LABVIEW_ISO%" (
    echo.
    echo LabVIEW ISO not found.
    echo   Expected location: %LABVIEW_ISO%
    echo.
    echo Options:
    echo   [1] Continue without LabVIEW (Visual Studio only^)
    echo   [2] Check again (after copying ISO^)
    echo   [3] Exit
    echo.
    choice /c 123 /n /m "Select option (1, 2, or 3): "
    if errorlevel 3 exit /b 0
    if errorlevel 2 goto :check_labview_iso
    goto :skip_labview_check
)

if exist "%SPECFILE_DIR%" goto :specfile_done

echo.
echo LabVIEW ISO found but no specfile exists.
echo   ISO: %LABVIEW_ISO%
echo   Expected specfile: %SPECFILE_DIR%
echo.
echo Options:
echo   [1] Continue without specfile (will prompt in sandbox^)
echo   [2] Create specfile now (launches NI installer GUI^)
echo   [3] Exit
echo.
choice /c 123 /n /m "Select option (1, 2, or 3): "
if errorlevel 3 exit /b 0
if errorlevel 2 goto :create_specfile
goto :specfile_done

:create_specfile
echo.
echo Mounting LabVIEW ISO...

:: Mount the ISO
powershell -Command "Mount-DiskImage -ImagePath '%LABVIEW_ISO%'" >nul 2>&1
ping -n 2 127.0.0.1 >nul

:: Find the mounted drive
set SPEC_DRIVE=
for %%d in (D E F G H I J K L) do (
    if exist "%%d:\setup.exe" (
        set SPEC_DRIVE=%%d
        goto :found_spec_drive
    )
)

:found_spec_drive
if not defined SPEC_DRIVE (
    echo ERROR: Could not find setup.exe on mounted ISO.
    echo.
    echo To generate a specfile manually:
    echo   1. Mount the LabVIEW ISO
    echo   2. Run: X:\setup.exe /generatespecfile "%SPECFILE_DIR%"
    echo   3. Select desired components in the installer GUI
    echo.
    pause
    goto :specfile_done
)

echo ISO mounted at drive %SPEC_DRIVE%:
echo.
echo Launching NI installer to generate specfile...
echo.
echo   *** SELECT YOUR DESIRED COMPONENTS IN THE INSTALLER GUI ***
echo   *** THEN CLOSE THE INSTALLER WHEN FINISHED ***
echo.
"%SPEC_DRIVE%:\setup.exe" /generatespecfile "%SPECFILE_DIR%"

:: Unmount the ISO
echo.
echo Unmounting ISO...
powershell -Command "Dismount-DiskImage -ImagePath '%LABVIEW_ISO%'" >nul 2>&1

if exist "%SPECFILE_DIR%" (
    echo Specfile created successfully!
) else (
    echo WARNING: Specfile was not created.
)

:specfile_done

:skip_labview_check
echo.
echo LabVIEW:
if exist "%LABVIEW_ISO%" (
    echo   ISO: Found
    if exist "%SPECFILE_DIR%" (
        echo   Specfile: Found
    ) else (
        echo   Specfile: Not found (will prompt in sandbox^)
    )
) else (
    echo   ISO: Not found (LabVIEW installation will be skipped^)
)

echo.
:: Generate the .wsb file with correct paths
echo Generating sandbox configuration...
echo   Shared folder: %SHARED_DIR%
echo   Output folder: %OUTPUT_DIR%

:: Start building the WSB file
(
echo ^<?xml version="1.0" encoding="UTF-8"?^>
echo ^<Configuration^>
echo   ^<VGpu^>Enable^</VGpu^>
echo   ^<Networking^>Enable^</Networking^>
echo   ^<MappedFolders^>
echo     ^<MappedFolder^>
echo       ^<HostFolder^>%SHARED_DIR%^</HostFolder^>
echo       ^<SandboxFolder^>C:\Users\WDAGUtilityAccount\Desktop\Setup^</SandboxFolder^>
echo       ^<ReadOnly^>true^</ReadOnly^>
echo     ^</MappedFolder^>
echo     ^<MappedFolder^>
echo       ^<HostFolder^>%OUTPUT_DIR%^</HostFolder^>
echo       ^<SandboxFolder^>C:\Users\WDAGUtilityAccount\Desktop\Output^</SandboxFolder^>
echo       ^<ReadOnly^>false^</ReadOnly^>
echo     ^</MappedFolder^>
) > "%WSB_FILE%"

:: Only add specfile mapping if the folder exists
if exist "%SPECFILE_DIR%" (
    echo   Specfile folder: %SPECFILE_DIR%
    (
echo     ^<MappedFolder^>
echo       ^<HostFolder^>%SPECFILE_DIR%^</HostFolder^>
echo       ^<SandboxFolder^>C:\Users\WDAGUtilityAccount\Desktop\specfile^</SandboxFolder^>
echo       ^<ReadOnly^>true^</ReadOnly^>
echo     ^</MappedFolder^>
    ) >> "%WSB_FILE%"
) else (
    echo   Specfile folder not found, skipping: %SPECFILE_DIR%
)

:: Finish the WSB file
(
echo   ^</MappedFolders^>
echo   ^<LogonCommand^>
echo     ^<Command^>cmd /c start cmd /k C:\Users\WDAGUtilityAccount\Desktop\Setup\install_all.bat^</Command^>
echo   ^</LogonCommand^>
echo   ^<MemoryInMB^>8192^</MemoryInMB^>
echo ^</Configuration^>
) >> "%WSB_FILE%"

:: Create desktop shortcut to the .wsb file
set SHORTCUT_PATH=%USERPROFILE%\Desktop\VS Sandbox.lnk
echo.
echo Creating desktop shortcut: VS Sandbox
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%SHORTCUT_PATH%'); $s.TargetPath = '%WSB_FILE%'; $s.Description = 'Launch Visual Studio Development Sandbox'; $s.Save()"

echo.
echo ============================================================================
echo  Setup Complete!
echo ============================================================================
echo.
echo Desktop shortcut created: VS Sandbox
echo.
echo You can now launch the sandbox by:
echo   - Double-clicking "VS Sandbox" on your desktop
echo   - Or double-clicking: %WSB_FILE%
echo.
pause
