@echo off
set sumber=C:\FolderAsal
set tujuan=%~dp0

robocopy "%sumber%" "%tujuan%" /E /R:0 /W:0