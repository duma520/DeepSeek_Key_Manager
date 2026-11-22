@echo off
rem ============================================================
rem  DeepSeek Key Manager - Nuitka build script (PySide6 / clang)
rem ------------------------------------------------------------
rem  NOTE: keep every comment in this file ASCII-only and keep
rem  the file saved as CRLF without BOM. cmd.exe parses the file
rem  byte by byte, so UTF-8 comments would break the parser.
rem
rem  Output: build_output\DeepSeek_Key_Manager.dist\
rem  Distribute the WHOLE dist folder (exe + icon.ico + DLLs).
rem
rem  Set EM_SKIP_BUILD=1 to only validate prerequisites (dry run).
rem ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo  DeepSeek Key Manager - Nuitka build
echo  Working directory: %CD%
echo ============================================================
echo.

if not exist "DeepSeek_Key_Manager.py" (
    echo [ERROR] DeepSeek_Key_Manager.py not found.
    echo         Put this bat next to the main script and run it again.
    exit /b 1
)

if not exist "icon.ico" (
    echo [WARN] icon.ico not found - the exe will have no icon and the
    echo        window/taskbar icon will be missing too.
    echo.
)

python -c "import PySide6, requests" 2>nul
if errorlevel 1 (
    echo [ERROR] Missing dependencies. Install them first:
    echo         pip install -r requirements.txt
    exit /b 1
)

echo [INFO] Prerequisites OK.
if defined EM_SKIP_BUILD (
    echo [DRY RUN] EM_SKIP_BUILD is set - stopping before the actual build.
    exit /b 0
)
echo.

echo [INFO] Building with Nuitka... this can take several minutes.
echo.

rem  Exclude heavy packages that this program never uses. Without these
rem  exclusions the dist folder balloons to ~430 MB because Nuitka also
rem  follows the OPTIONAL imports found inside openpyxl / pandas / etc.
rem  Keep this list in sync when new third-party imports are added.
set NUITKA_EXCLUDES=matplotlib,numpy,pandas,scipy,contourpy,kiwisolver,pyparsing,cycler,PIL,Pillow,tkinter,IPython,jupyter,jupyter_core,cryptography,lxml,pythonnet,clr_loader,pytz,markupsafe,backports,setuptools,pip,wheel,pkg_resources,test,unittest,doctest,lib2to3,distutils,venv,ensurepip

python -m nuitka ^
    --standalone ^
    --assume-yes-for-downloads ^
    --enable-plugin=pyside6 ^
    --windows-console-mode=disable ^
    --windows-icon-from-ico=icon.ico ^
    --include-data-files=icon.ico=icon.ico ^
    --include-package=requests ^
    --include-package=openpyxl ^
    --include-package=certifi ^
    --include-package-data=certifi ^
    --include-package=pypinyin ^
    --include-package-data=pypinyin ^
    --nofollow-import-to=%NUITKA_EXCLUDES% ^
    --follow-imports ^
    --jobs=4 ^
    --clang ^
    --remove-output ^
    --output-dir=build_output ^
    DeepSeek_Key_Manager.py

set NUITKA_EXIT=%ERRORLEVEL%
set NUITKA_EXCLUDES=
if not "%NUITKA_EXIT%"=="0" (
    echo.
    echo [ERROR] Nuitka build failed with exit code %NUITKA_EXIT%. Common reasons:
    echo         - an old DeepSeek_Key_Manager.exe is still running
    echo           (close it and retry; the linker cannot overwrite it)
    echo         - missing clang / MSVC toolchain
    exit /b 1
)

echo.
echo ============================================================
echo  Build finished.
echo  Output: %CD%\build_output\DeepSeek_Key_Manager.dist\
echo  Remember: distribute the whole .dist folder, and keep the
echo  user's .env file OUT of the release package.
echo ============================================================
