@echo off
setlocal enabledelayedexpansion
echo Removendo comandos no registro...


set current_directory=%~dp0
set icon_filename=ico.ico
set icon_filepath=%current_directory%%icon_filename%
set exe_filename=exe.exe
set exe_filepath=%current_directory%%exe_filename%

set "files=png jpg jpeg webp ico tiff"
for %%a in (%files%) do (
  reg delete "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage" /f
)


echo Comandos removidos.
pause