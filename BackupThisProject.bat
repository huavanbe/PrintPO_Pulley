@echo off
FOR /F %%i IN ('cd') DO set ProjectName=%%~nxi
set /p BackupFolder=<"D:\Be\Projects\BackupPath.txt"
for %%I in (.) do set ProjectFolder=%%~nI%%~xI
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /format:list') do set datetime=%%I
"C:\Program Files\7-Zip\7z.exe"  a "%BackupFolder%\%ProjectName%\%ProjectName%_%datetime:~0,8%_%datetime:~8,4%.zip" ..\%ProjectFolder%\ -xr!packages

::WinRAR a help *.hlp
@pause


