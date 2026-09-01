@echo off
setlocal
rem Opens the system in Google Chrome specifically (not the default browser),
rem because Chrome supports IndexedDB from a local file, while some other
rem browsers block it. Right-click this file > "Send to" > "Desktop
rem (create shortcut)" and rename the shortcut to "نظام المبيعات" to get a
rem double-click icon.
start "" chrome "%~dp0index.html"
