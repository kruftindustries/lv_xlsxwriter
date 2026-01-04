@echo off
setlocal EnableDelayedExpansion

:: ============================================================================
:: libxlsxwriter Build Script for Windows
:: Builds zlib and libxlsxwriter for both 32-bit and 64-bit
:: ============================================================================

set DESKTOP=%USERPROFILE%\Desktop
set SETUP_DIR=%DESKTOP%\Setup
set ZLIB_SRC=%DESKTOP%\zlib
set XLSXWRITER_SHARED=%SETUP_DIR%\libxlsxwriter-1.2.3
set XLSXWRITER_SRC=%DESKTOP%\libxlsxwriter-1.2.3
set VS_VERSION=Visual Studio 18 2026

echo ============================================================================
echo  libxlsxwriter Build Script
echo  Started: %DATE% %TIME%
echo ============================================================================
echo.
echo Source directories:
echo   zlib:        %ZLIB_SRC%
echo   xlsxwriter:  %XLSXWRITER_SRC% (copied from shared)
echo.

:: ============================================================================
:: Step 1: Clone zlib and copy libxlsxwriter to local directory
:: ============================================================================
echo [Step 1/6] Setting up source directories...
echo.

:: Clone zlib if not present
if exist "%ZLIB_SRC%" (
    echo zlib already exists, skipping clone.
) else (
    echo Cloning zlib...
    cd /d "%DESKTOP%"
    git clone https://github.com/madler/zlib.git
    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Failed to clone zlib!
        pause
        exit /b 1
    )
)

:: Copy libxlsxwriter from shared folder to local directory
:: (Mapped folders have issues with CMake)
echo Copying libxlsxwriter to local directory...
if exist "%XLSXWRITER_SRC%" rmdir /s /q "%XLSXWRITER_SRC%"
xcopy "%XLSXWRITER_SHARED%" "%XLSXWRITER_SRC%\" /s /e /q /i
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy libxlsxwriter!
    pause
    exit /b 1
)
echo.

:: ============================================================================
:: Step 2: Build zlib 32-bit
:: ============================================================================
echo [Step 2/6] Building zlib 32-bit...
echo.

::call "%ProgramFiles%\Microsoft Visual Studio\2026\Community\VC\Auxiliary\Build\vcvars32.bat"
call "%ProgramFiles%\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars32.bat"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to set up 32-bit environment!
    pause
    exit /b 1
)

cd /d "%ZLIB_SRC%"
if exist build32 rmdir /s /q build32
mkdir build32
cd build32

cmake .. -G "%VS_VERSION%" -A Win32 -DBUILD_SHARED_LIBS=OFF -DCMAKE_MSVC_RUNTIME_LIBRARY=MultiThreadedDLL
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: CMake configure failed for zlib 32-bit!
    pause
    exit /b 1
)

cmake --build . --config Release
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed for zlib 32-bit!
    pause
    exit /b 1
)

echo zlib 32-bit build complete.
echo.

:: ============================================================================
:: Step 3: Build libxlsxwriter 32-bit
:: ============================================================================
echo [Step 3/6] Building libxlsxwriter 32-bit...
echo.

cd /d "%XLSXWRITER_SRC%"
if exist build32 rmdir /s /q build32
mkdir build32
cd build32

cmake .. -G "%VS_VERSION%" -A Win32 ^
    -DZLIB_ROOT="%ZLIB_SRC%\build32\Release" ^
    -DZLIB_LIBRARY="%ZLIB_SRC%\build32\Release\zs.lib" ^
    -DZLIB_INCLUDE_DIR="%ZLIB_SRC%" ^
    -DBUILD_SHARED_LIBS=ON ^
    -DUSE_STATIC_MSVC_RUNTIME=OFF ^
    -DCMAKE_C_BYTE_ORDER=LITTLE_ENDIAN ^
    -DCMAKE_MSVC_RUNTIME_LIBRARY=MultiThreadedDLL

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: CMake configure failed for libxlsxwriter 32-bit!
    pause
    exit /b 1
)

cmake --build . --config Release
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed for libxlsxwriter 32-bit!
    pause
    exit /b 1
)

echo libxlsxwriter 32-bit build complete.
echo.

:: ============================================================================
:: Step 4: Build zlib 64-bit
:: ============================================================================
echo [Step 4/6] Building zlib 64-bit...
echo.

::call "%ProgramFiles%\Microsoft Visual Studio\2026\Community\VC\Auxiliary\Build\vcvars64.bat"
call "%ProgramFiles%\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars64.bat"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to set up 64-bit environment!
    pause
    exit /b 1
)

cd /d "%ZLIB_SRC%"
if exist build64 rmdir /s /q build64
mkdir build64
cd build64

cmake .. -G "%VS_VERSION%" -A x64 -DBUILD_SHARED_LIBS=OFF -DCMAKE_MSVC_RUNTIME_LIBRARY=MultiThreadedDLL
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: CMake configure failed for zlib 64-bit!
    pause
    exit /b 1
)

cmake --build . --config Release
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed for zlib 64-bit!
    pause
    exit /b 1
)

echo zlib 64-bit build complete.
echo.

:: ============================================================================
:: Step 5: Build libxlsxwriter 64-bit
:: ============================================================================
echo [Step 5/6] Building libxlsxwriter 64-bit...
echo.

cd /d "%XLSXWRITER_SRC%"
if exist build64 rmdir /s /q build64
mkdir build64
cd build64

cmake .. -G "%VS_VERSION%" -A x64 ^
    -DZLIB_ROOT="%ZLIB_SRC%\build64\Release" ^
    -DZLIB_LIBRARY="%ZLIB_SRC%\build64\Release\zs.lib" ^
    -DZLIB_INCLUDE_DIR="%ZLIB_SRC%" ^
    -DBUILD_SHARED_LIBS=ON ^
    -DUSE_STATIC_MSVC_RUNTIME=OFF ^
    -DCMAKE_C_BYTE_ORDER=LITTLE_ENDIAN ^
    -DCMAKE_MSVC_RUNTIME_LIBRARY=MultiThreadedDLL

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: CMake configure failed for libxlsxwriter 64-bit!
    pause
    exit /b 1
)

cmake --build . --config Release
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Build failed for libxlsxwriter 64-bit!
    pause
    exit /b 1
)

echo libxlsxwriter 64-bit build complete.
echo.

:: ============================================================================
:: Step 6: Package output files
:: ============================================================================
echo [Step 6/6] Packaging output files...
echo.

set OUTPUT_DIR=%DESKTOP%\Output
set PKG_DIR=%OUTPUT_DIR%\pkg
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

:: Package 32-bit
echo Creating libxlsxwriter-1.2.4-x86.zip...
set PKG32=%PKG_DIR%\libxlsxwriter-1.2.4-x86
if exist "%PKG32%" rmdir /s /q "%PKG32%"
mkdir "%PKG32%"
mkdir "%PKG32%\lib"
mkdir "%PKG32%\include"

copy "%XLSXWRITER_SRC%\build32\Release\xlsxwriter.dll" "%PKG32%\lib\" >nul
copy "%XLSXWRITER_SRC%\build32\Release\xlsxwriter.lib" "%PKG32%\lib\" >nul
xcopy "%XLSXWRITER_SRC%\include\*.h" "%PKG32%\include\" /s /q >nul

powershell -NoProfile -Command "Compress-Archive -Path '%PKG32%\*' -DestinationPath '%OUTPUT_DIR%\libxlsxwriter-1.2.4-x86.zip' -Force"
if %ERRORLEVEL% NEQ 0 (
    echo WARNING: Failed to create x86 zip file
) else (
    echo Created: %OUTPUT_DIR%\libxlsxwriter-1.2.4-x86.zip
)

:: Package 64-bit
echo Creating libxlsxwriter-1.2.4-x64.zip...
set PKG64=%PKG_DIR%\libxlsxwriter-1.2.4-x64
if exist "%PKG64%" rmdir /s /q "%PKG64%"
mkdir "%PKG64%"
mkdir "%PKG64%\lib"
mkdir "%PKG64%\include"

copy "%XLSXWRITER_SRC%\build64\Release\xlsxwriter.dll" "%PKG64%\lib\" >nul
copy "%XLSXWRITER_SRC%\build64\Release\xlsxwriter.lib" "%PKG64%\lib\" >nul
xcopy "%XLSXWRITER_SRC%\include\*.h" "%PKG64%\include\" /s /q >nul

powershell -NoProfile -Command "Compress-Archive -Path '%PKG64%\*' -DestinationPath '%OUTPUT_DIR%\libxlsxwriter-1.2.4-x64.zip' -Force"
if %ERRORLEVEL% NEQ 0 (
    echo WARNING: Failed to create x64 zip file
) else (
    echo Created: %OUTPUT_DIR%\libxlsxwriter-1.2.4-x64.zip
)

:: Clean up temporary package directories
rmdir /s /q "%PKG_DIR%" 2>nul

echo.

:: ============================================================================
:: Summary
:: ============================================================================
echo ============================================================================
echo  Build Complete!
echo  Finished: %DATE% %TIME%
echo ============================================================================
echo.
echo Output packages:
echo   %OUTPUT_DIR%\libxlsxwriter-1.2.4-x86.zip
echo   %OUTPUT_DIR%\libxlsxwriter-1.2.4-x64.zip
echo.
echo Package contents:
echo   lib\xlsxwriter.dll  - Dynamic library
echo   lib\xlsxwriter.lib  - Import library
echo   include\            - Header files
echo.

pause
