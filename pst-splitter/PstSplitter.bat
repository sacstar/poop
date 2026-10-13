@echo off
rem Opens the PST Splitter window. Outlook must be installed on this PC.
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0PstSplitter.ps1" %*
