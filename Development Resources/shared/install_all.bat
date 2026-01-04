@echo off
setlocal EnableDelayedExpansion

:: Apply dark theme without waiting for user interaction
start "" /min C:\Windows\Resources\Themes\dark.theme
ping -n 2 127.0.0.1 >nul
taskkill /IM systemsettings.exe /F >nul 2>&1

:: ============================================================================
:: Development Environment Installer for Windows Sandbox
:: Installs Visual Studio Community 2026 with C++ and LabVIEW
:: ============================================================================

:: Mapped folder from host (read-only)
set SETUP_DIR=%USERPROFILE%\Desktop\Setup
:: Output folder mapped from host (writable) - for build output
set OUTPUT_DIR=%USERPROFILE%\Desktop\Output
:: Local sandbox folder for temporary files
set LOCAL_DIR=%USERPROFILE%\Desktop\Local
set LOG_FILE=%OUTPUT_DIR%\install_log.txt
set LOG_DIR=%OUTPUT_DIR%\logs
mkdir "%LOCAL_DIR%" 2>nul
mkdir "%LOG_DIR%" 2>nul

:: ============================================================================
:: Enable logging - redirect all output to both console and log file
:: ============================================================================
:: Note: PowerShell Tee-Object logging disabled because it breaks interactive
:: prompts (choice command). Using simple file logging instead.
:: if "%LOGGING_ENABLED%"=="" (
::     set LOGGING_ENABLED=1
::     echo Logging to: %LOG_FILE%
::     powershell -Command "& {cmd /c 'set LOGGING_ENABLED=1 && \"%~f0\"'} 2>&1 | Tee-Object -FilePath '%LOG_FILE%'"
::     exit /b %ERRORLEVEL%
:: )

:: Start logging to file (append mode, doesn't interfere with interactive input)
echo. > "%LOG_FILE%"
echo Installation started: %DATE% %TIME% >> "%LOG_FILE%"

echo ============================================================================
echo  Development Environment Installer
echo  Started: %DATE% %TIME%
echo ============================================================================
echo.
echo Setup Directory: %SETUP_DIR%
echo Log File: %LOG_FILE%
echo.

:: ============================================================================
:: Step 0: Apply Windows Sandbox Performance Workaround
:: ============================================================================
:: Disables Code Integrity policy verification which causes slow execution
:: in Windows Sandbox. This significantly speeds up installer execution.
:: ============================================================================
echo [Step 0/5] Applying Windows Sandbox performance optimization...
echo.

powershell -Command "Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy' -Name 'VerifiedAndReputablePolicyState' -Value 0" 2>nul
echo. | CiTool.exe -r >nul 2>&1

echo Performance optimization applied.
echo.

:: ============================================================================
:: Step 1: Check for Offline Layout or Download Installer
:: ============================================================================
set VS_LAYOUT=%SETUP_DIR%\VSLayout\vs_community.exe
set VS_INSTALLER=%LOCAL_DIR%\vs_community.exe

echo Checking for VS layout at: %VS_LAYOUT%

if exist "%VS_LAYOUT%" goto :use_vs_layout

echo [Step 1/5] Downloading Visual Studio Community Installer...
echo   (No offline layout found at %VS_LAYOUT%)
echo.
curl -L -o "%VS_INSTALLER%" --progress-bar "https://aka.ms/vs/stable/vs_community.exe"
if not exist "%VS_INSTALLER%" (
    echo ERROR: Failed to download Visual Studio installer!
    pause
    exit /b 1
)
echo Download complete.
set USE_LAYOUT=0
goto :vs_check_done

:use_vs_layout
echo [Step 1/5] Using offline Visual Studio layout...
echo   Layout: %VS_LAYOUT%
set VS_INSTALLER=%VS_LAYOUT%
set USE_LAYOUT=1

:vs_check_done
echo.

:: ============================================================================
:: Step 2: Install Visual Studio with C++ Desktop Development
:: ============================================================================
echo [Step 2/5] Installing Visual Studio Community...
echo            - C++ Desktop Development workload
echo            - CMake tools, Windows SDK
echo.
if "%USE_LAYOUT%"=="1" (
    echo This may take 5-15 minutes. Please wait...
)
echo.

:: Note: --log option causes error 87 in sandbox, so omitted
echo VS Installer: %VS_INSTALLER%
echo.

if "%USE_LAYOUT%"=="1" (
    :: Layout install - components already selected during layout creation
    "%VS_INSTALLER%" --wait --passive --norestart
) else (
    :: Online install - specify components
    "%VS_INSTALLER%" --wait --passive --norestart --add Microsoft.VisualStudio.Workload.NativeDesktop --add Microsoft.VisualStudio.Component.VC.Tools.x86.x64 --add Microsoft.VisualStudio.Component.Windows11SDK.22621 --add Microsoft.VisualStudio.Component.VC.CMake.Project --add Microsoft.VisualStudio.Component.Git --remove Component.GitHub.Copilot --remove Component.GitHub.Copilot.Chat --includeRecommended
)

if %ERRORLEVEL% NEQ 0 (
    if %ERRORLEVEL% NEQ 3010 (
        echo WARNING: Visual Studio installation returned code %ERRORLEVEL%
    )
)

echo Visual Studio installation complete.
echo Visual Studio installation complete: %DATE% %TIME% >> "%LOG_FILE%"
echo.

:: ============================================================================
:: Step 3: Install LabVIEW from ISO
:: ============================================================================
echo [Step 3/5] Installing LabVIEW...
echo.

set LABVIEW_ISO=%SETUP_DIR%\2015myRIO Software Bundle 1.iso
:: Check for specfile in host shared folder first, then local
if exist "%SETUP_DIR%\specfile_lv" (
    set LABVIEW_SPECFILE=%SETUP_DIR%\specfile_lv
) else (
    set LABVIEW_SPECFILE=%LOCAL_DIR%\specfile
)

if not exist "%LABVIEW_ISO%" (
    echo WARNING: LabVIEW ISO not found.
    echo   Expected location: %LABVIEW_ISO%
    echo.
    choice /c YN /n /m "Continue without LabVIEW installation? (Y/N): " < CON
    if errorlevel 2 (
        echo.
        echo Please place the LabVIEW ISO at: %LABVIEW_ISO%
        echo Then restart this script.
        pause
        exit /b 1
    )
    echo.
    echo Skipping LabVIEW installation.
    goto :skip_labview
)

if not exist "%LABVIEW_SPECFILE%" (
    echo WARNING: LabVIEW specfile not found.
    echo   Expected: %LABVIEW_SPECFILE%
    echo.
    echo To create a specfile, run launch_dev_sandbox.bat on the host.
    echo.
    echo Skipping LabVIEW installation.
    goto :skip_labview
)

:: Mount the ISO using PowerShell and detect the drive letter
echo Mounting LabVIEW ISO...
echo   ISO: %LABVIEW_ISO%

:: Clear any previous value
set DRIVE_LETTER=

:: Use PowerShell to mount the ISO silently
powershell -Command "Mount-DiskImage -ImagePath '%LABVIEW_ISO%'" >nul 2>&1
ping -n 2 127.0.0.1 >nul

:: Find the mounted drive by checking for setup.exe on likely drive letters
for %%d in (D E F G H I J K L) do (
    if exist "%%d:\setup.exe" (
        set DRIVE_LETTER=%%d
        goto :found_drive
    )
)

:found_drive
if not defined DRIVE_LETTER (
    echo ERROR: Failed to find mounted ISO drive!
    echo   Checked drives D-L for setup.exe
    goto :skip_labview
)

echo ISO mounted at drive %DRIVE_LETTER%:
echo.

:: Check for setup executable
set LABVIEW_SETUP=%DRIVE_LETTER%:\setup.exe
if not exist "%LABVIEW_SETUP%" (
    echo ERROR: Could not find setup.exe on the mounted ISO!
    goto :unmount_iso
)

:: Run LabVIEW installer with specfile
echo Running LabVIEW setup with specfile...
echo   Specfile: %LABVIEW_SPECFILE%
echo.
echo This may take 5-25 minutes. Please wait...
echo.

"%LABVIEW_SETUP%" /applyspecfile "%LABVIEW_SPECFILE%" /q /acceptlicenses yes /r:n /disableNotificationCheck

if %ERRORLEVEL% NEQ 0 (
    echo WARNING: LabVIEW installation returned code %ERRORLEVEL%
)

echo LabVIEW installation complete.
echo LabVIEW installation complete: %DATE% %TIME% >> "%LOG_FILE%"

:unmount_iso
:: Unmount the ISO
echo.
echo Unmounting ISO...
powershell -Command "Dismount-DiskImage -ImagePath '%LABVIEW_ISO%'"

:skip_labview
echo.

:: ============================================================================
:: Step 4: Setup Environment Variables
:: ============================================================================
echo [Step 4/5] Configuring environment...
echo.

:: Set VS environment variables - check for script path, installed for "Desktop development with C++ workload"
:: [Edition] can be Community, Professional, or Enterprise. 
:: C:\Program Files\Microsoft Visual Studio\[Version]\[Edition]\VC\Auxiliary\Build\vcvars[CPU Architecture].bat
:: example for VS 2022:
:: C:\Program Files\Microsoft Visual Studio\2022\[Edition]\VC\Auxiliary\Build\
:: for VS 2017/2019
:: C:\Program Files (x86)\Microsoft Visual Studio\2019\[Edition]\VC\Auxiliary\Build\
set VS_VCVARS=
if exist "%ProgramFiles%\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars32.bat" (
    set VS_VCVARS=%ProgramFiles%\Microsoft Visual Studio\2026\Community\VC\Auxiliary\Build\vcvars[32/64].bat
) else if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars32.bat" (
    set VS_VCVARS=%ProgramFiles%\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars[32/64].bat
)

:: if defined VS_VCVARS (
::     call "%VS_VCVARS%" 2>nul
:: )

:: ============================================================================
:: Step 5: Open Command Prompt with VS Environment
:: ============================================================================
echo [Step 5/5] Opening Developer Command Prompt...
echo.

echo.
echo ============================================================================
echo  Installation Complete!
echo  Finished: %DATE% %TIME%
echo ============================================================================
echo Installation complete: %DATE% %TIME% >> "%LOG_FILE%"
echo.
echo Visual Studio Community has been installed with:
echo   - C++ Desktop Development workload
echo   - CMake tools
echo   - Windows 11 SDK
echo.
if defined VS_VCVARS (
    echo To use the Visual C++ environment, run:
    echo   "%VS_VCVARS%"
) else (
    echo Visual Studio environment script not found.
    echo Please locate vcvars64.bat manually.
)
echo.
echo To build libxlsxwriter, run:
echo   %SETUP_DIR%\build.bat
echo.
echo Log files:
echo   - Main log: %LOG_FILE%
echo.

:: Check if we should auto-run build
if exist "%SETUP_DIR%\AUTO_BUILD" (
    echo.
    echo AUTO_BUILD flag detected, starting build...
    echo.
    call "%SETUP_DIR%\build.bat"

    :: Signal completion
    echo DONE > "%SETUP_DIR%\BUILD_COMPLETE"
    echo.
    echo Build complete. Signaled completion via BUILD_COMPLETE file.
    echo Sandbox will close in 10 seconds...
    ping -n 11 127.0.0.1 >nul
    exit
) else (
    :: Keep the window open for interactive use
    cmd /k
)
