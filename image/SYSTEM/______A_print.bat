@echo off
setlocal

:: 現在の作業ディレクトリを取得
set "WORK_DIR=%~dp0"

:: レジストリからPhotoshopのインストールパスを取得
for /f "tokens=2,*" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe" /v "" 2^>nul') do set "PHOTOSHOP_PATH=%%B"

:: レジストリからPhotoshopのインストールパスが取得できなかった場合、エラーメッセージを表示して終了
if "%PHOTOSHOP_PATH%"=="" (
    echo Photoshopのインストールパスがレジストリから取得できませんでした。
    exit /b 1
)

:: JSXスクリプトの相対パスをフルパスに変換
set "REL_SCRIPT_PATH=..\image\SYSTEM\printerA.jsx"
call :ResolvePath "%WORK_DIR%%REL_SCRIPT_PATH%" SCRIPT_PATH

:: ARGUMENT_PATHの初期化（完全に空のファイルを作成）
set "REL_ARGUMENT_PATH=..\image\SYSTEM\argument.info"
call :ResolvePath "%WORK_DIR%%REL_ARGUMENT_PATH%" ARGUMENT_PATH
copy nul "%ARGUMENT_PATH%" >nul 2>&1

:: ドラッグされた各ファイルの処理
:ProcessFiles
if "%~1"=="" goto :RunPhotoshop

set "FILE_PATH=%~1"

:: パスの一部を置き換える
set "FILE_PATH_MODIFIED=%FILE_PATH:focetta マクロ=focettaマクロ%"

:: 一時ファイルにPSDファイルパスを保存
>> "%ARGUMENT_PATH%" echo %FILE_PATH_MODIFIED%

shift
goto :ProcessFiles

:RunPhotoshop
:: Photoshopをバックグラウンドで実行
echo start "" "%PHOTOSHOP_PATH%" -r "%SCRIPT_PATH%"
start "" "%PHOTOSHOP_PATH%" -r "%SCRIPT_PATH%"

endlocal
exit /b

:ResolvePath
:: フルパスを解決するためのサブルーチン
setlocal
set "FULL_PATH=%~f1"
endlocal & set "%2=%FULL_PATH%"
goto :eof
