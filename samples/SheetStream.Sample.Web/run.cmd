@echo off
setlocal
rem Double-click to run the SheetStream web sample. Builds, starts the server on a fixed port and opens the
rem browser. The window stays open on failure so the message can be read.
cd /d "%~dp0"
set URL=http://localhost:5178

echo Building the SheetStream web sample (first time takes a little longer)...
dotnet build -c Release --nologo -v quiet
if errorlevel 1 goto failed

echo.
echo Starting the server at %URL%
echo Your browser opens in a moment. Close this window or press Ctrl+C to stop the server.
echo.
rem Open the browser a few seconds from now, by which time the server is listening.
start "" /min cmd /c "timeout /t 3 /nobreak >nul && start "" %URL%"
dotnet run -c Release --no-build --urls %URL%
if errorlevel 1 goto failed
exit /b 0

:failed
echo.
echo The sample did not start. Read the message above. The usual causes are:
echo   - the .NET 8 SDK is not installed (download it from https://dotnet.microsoft.com/download)
echo   - another program is already using port 5178
pause
exit /b 1
