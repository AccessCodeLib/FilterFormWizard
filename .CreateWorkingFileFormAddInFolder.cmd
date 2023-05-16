@echo off

if exist .\ACLibFilterFormWizard.accdb (
set /p CopyFile=ACLibFilterFormWizard.accdb exists .. overwrite with access-add-in\ACLibFilterFormWizard.accda? [Y/N]:
) else (
set CopyFile=Y
)

if /I %CopyFile% == Y (
	echo File is copied ...
) else (
	echo Batch is cancelled
	pause
	exit
)

copy .\access-add-in\ACLibFilterFormWizard.accda ACLibFilterFormWizard.accdb

timeout 2