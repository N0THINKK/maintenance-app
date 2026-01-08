@echo OFF

ren C:\AC90HMI\prg\PrdLog.csv PrdLog-%date:~10,4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%.csv> PrdLog-%date:~10,4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%.csv

xcopy "C:\Paperless\PrdLog.csv" "C:\AC90HMI\prg" /H /C /Y