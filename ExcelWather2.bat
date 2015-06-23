@echo off
Set fileName=LCHCME.xls
Set fileNameHelper=ExcelHelper.xls
Set excelProgramPath="C:\Program Files (x86)\Microsoft Office\Office14\excel.exe"
set retryLimits=3
set retryAttempts=0
set isOpen=0


:Check
SETLOCAL enabledelayedexpansion 
::First check if excel file is being opened or not
start /MIN "" %excelProgramPath% %fileNameHelper%
cscript ExcelHelper.vbs delay 1 //nologo
for /F "delims=" %%i in (temp.txt) do set "isOpen=%%i"

::File is not opened
if %isOpen% EQU 0 (
	ECHO %date% %time% : %fileName% is not opened.
	ECHO %date% %time% : Now trying to open %fileName%.
	set /A retryAttempts=%retryAttempts%+1
	start /MIN "" %excelProgramPath% %fileName%
	cscript ExcelHelper.vbs delay 3 //nologo
	ExcelWather2.bat
) 
::File is opened
if %isOpen% EQU 1 (
	ECHO %date% %time% : %fileName% has already been opened.
	ECHO %date% %time% : Now monitor the log file.
	cscript ExcelHelper.vbs delay 1 //nologo
)
endlocal

:Monitor
SETLOCAL enabledelayedexpansion 
for /F "delims=" %%i in (C:\Users\xili\Documents\log.txt) do set "lastLine=%%i"
for /F "tokens=2 delims= " %%i in ("%lastLine%") do set "lastLogTime=%%i"
echo %date% %time% : Last Updated Time: %lastLogTime%.


set currentTime=%TIME%
::adjust the time format
for /F "tokens=1 delims=:/ " %%i in ("%currentTime%") do (
	if %%i LSS 10 set currentTime=%currentTime: =0%
)

::calculate the days
for /F "tokens=2 delims=/" %%i in ("%lastLine%") do set "lastLogDay=%%i"
for /f "tokens=3 delims=/ " %%i in ('date /t') do set "currentDay=%%i"

set /A days=%currentDay%-%lastLogDay%

IF %days% LSS 0 set /A days=0
set interval=3600 


set /A lastLogTime=(1%lastLogTime:~0,2%-100)*3600+(1%lastLogTime:~3,2%-100)*60+(1%lastLogTime:~6,2%-100)

set /A currentTime=(1%currentTime:~0,2%-100)*3600+(1%currentTime:~3,2%-100)*60+(1%currentTime:~6,2%-100)
::ECHO %days%
::calculating duration (in seconds)
set /A duration=%currentTime%-%lastLogTime%+(%days%*24*60*60)
if %currentTime% LSS %lastLogTime% set /A duration=%lastLogTime%-%currentTime%

::now break the seconds down to hours, minutes
set /A durationH=%duration% / 3600
set /A durationM=(%duration% - %durationH%*3600) / 60
set /A durationS=(%duration% - %durationH%*3600 - %durationM%*60)

set /A intervalH=%interval% / 3600
set /A intervalM=(%interval% - %intervalH%*3600) / 60
set /A intervalS=(%interval% - %intervalH%*3600 - %intervalM%*60)

if %durationH% LSS 10 set durationH=0%durationH%
if %durationM% LSS 10 set durationM=0%durationM%
if %durationS% LSS 10 set durationS=0%durationS%

if %intervalH% LSS 10 set intervalH=0%intervalH%
if %intervalM% LSS 10 set intervalM=0%intervalM%
if %intervalS% LSS 10 set intervalS=0%intervalS%

echo %date% %time% : It has been %durationH%:%durationM%:%durationS% since last update.
ECHO %date% %time% : The alert interval time: %intervalH%:%intervalM%:%intervalS%

::
if %duration% GTR %interval% (
	ECHO %date% %time% : The excel file may not work normally.	
	ECHO %date% %time% : Now trying to reopen %fileName%.
	set /A retryAttempts=%retryAttempts%+1
	if %retryAttempts% LSS %retryLimits% (
		cscript ExcelHelper.vbs CloseExcel %fileName% //nologo
		cscript ExcelHelper.vbs delay 3 //nologo
		start /MIN "" %excelProgramPath% %fileName%
		cscript ExcelHelper.vbs delay 3 //nologo
		Call :Check
	) else (
		ECHO %date% %time% : Fatal Error.
		cscript ExcelHelper.vbs EmailSender "Fatal Error. Retry limit exceeded." //nologo
	)
) else (
	ECHO %date% %time% : Everything is fine.
	ECHO %date% %time% : The next check will be in 2 minutes
	cscript ExcelHelper.vbs delay 120 //nologo
	ExcelWather2.bat
)
endlocal
