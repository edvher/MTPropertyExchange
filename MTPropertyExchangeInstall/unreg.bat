@echo off
:: Thin wrapper - all registration logic lives in install.bat so that the
:: 32-bit and 64-bit registration is done in exactly one place.
::
::   unreg.bat                ... unregister COM components
::   unreg.bat /unregister    ... unregister COM components
::   unreg.bat /register      ... (re-)register COM components
::   unreg.bat /reg           ... (re-)register COM components
pushd "%~dp0"
if "%1"=="" call install.bat /unregister
if "%1"=="/unregister" call install.bat /unregister
if "%1"=="/register" call install.bat /register
if "%1"=="/reg" call install.bat /register
popd
