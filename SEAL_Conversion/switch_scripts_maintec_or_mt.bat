@echo off
pushd "%~dp0"
cls

echo.
if exist MT_property_update_office.pl goto activate_mt
if exist MAINTEC_property_update_office.pl goto activate_maintec
echo Nothing to do?
goto fine

:activate_mt
echo Currently, the MAINTEC script version (supports Office 2003 file formats only) is active:
echo.
dir /b *.pl
echo.
echo Activating MT script version (supports all Office file formats) ...
ren VAI_property_update_office.pl MAINTEC_property_update_office.pl
ren MT_property_update_office.pl VAI_property_update_office.pl 
goto fine

:activate_maintec
echo Currently, the MT script version (supports all Office file formats) is active:
echo.
dir /b *.pl
echo.
echo Activating MAINTEC script version (supports Office 2003 file formats only) ...
ren VAI_property_update_office.pl MT_property_update_office.pl
ren MAINTEC_property_update_office.pl VAI_property_update_office.pl 
goto fine

:fine
echo.
dir /b *.pl
echo.
echo Done.
echo.
pause