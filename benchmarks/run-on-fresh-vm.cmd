@echo off
rem Double-click launcher for run-on-fresh-vm.ps1. A fresh Windows machine does not allow .ps1 files to run by
rem default; this starts PowerShell with the permission needed for this one script only. Extra arguments are
rem passed on, for example:  run-on-fresh-vm.cmd -DryRun
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-on-fresh-vm.ps1" %*
if errorlevel 1 pause
