@echo off

if exist .\DbmsConnectionWizard.accdb (
set /p CopyFile=DbmsConnectionWizard.accdb exists .. overwrite with access-add-in\DbmsConnectionWizard.accda? [Y/N]:
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

copy .\access-add-in\DbmsConnectionWizard.accda DbmsConnectionWizard.accdb

timeout 2