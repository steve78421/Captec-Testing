@echo off
set "script=%~dp0masterScript.ps1"
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%script%""' -Verb RunAs}"
