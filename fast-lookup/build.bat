@echo off
rem Builds FastLookup.exe with the C# compiler built into Windows (.NET Framework 4). Nothing to install.
setlocal
cd /d "%~dp0"
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo Could not find csc.exe from .NET Framework 4.
  if not "%1"=="ci" pause
  exit /b 1
)
"%CSC%" /nologo /target:winexe /optimize+ /codepage:65001 /win32icon:app.ico /out:FastLookup.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll FastLookup.cs
if errorlevel 1 (
  echo Build failed.
  if not "%1"=="ci" pause
  exit /b 1
)
echo Built FastLookup.exe
if not "%1"=="ci" pause
