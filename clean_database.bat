@echo off
call "%~dp0start.bat" --purge %*
exit /b %ERRORLEVEL%
