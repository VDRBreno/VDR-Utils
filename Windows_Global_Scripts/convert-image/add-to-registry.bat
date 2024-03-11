@echo off
setlocal enabledelayedexpansion
echo Criando comando no registro...


set current_directory=%~dp0
set icon_filename=ico.ico
set icon_filepath=%current_directory%%icon_filename%
set exe_filename=exe.exe
set exe_filepath=%current_directory%%exe_filename%

set "files=png jpg jpeg webp ico tiff"
for %%a in (%files%) do (
  reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage" /f /ve /d ""
  reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage" /v "MUIVerb" /f /d "Converter Imagem"
  reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage" /v "Icon" /f /d %icon_filepath%

  set "commands="
  for %%b in (%files%) do (
    if not "%%a"=="%%b" (
      reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage\shell\%%b" /f /ve /d "%%b"
      reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage\shell\%%b\command" /f /ve /d "\"%exe_filepath%\" \"%%1\" \"%%b\""

      if "!commands!"=="" (
        set "commands=%%b"
      ) else (
        set "commands=!commands!;%%b"
      )
    )
  )

  reg add "HKEY_CLASSES_ROOT\SystemFileAssociations\.%%a\shell\ConvertImage" /v "SubCommands" /f /d "!commands!"
)


echo Comando criado no registro.
pause

@REM OLD-VERSION

@REM @echo off
@REM setlocal enabledelayedexpansion
@REM echo Criando comando no registro...


@REM reg add "HKEY_CLASSES_ROOT\*\shell\ConvertImage" /f /ve /d ""
@REM reg add "HKEY_CLASSES_ROOT\*\shell\ConvertImage" /v "MUIVerb" /f /d "Converter Imagem"
@REM reg add "HKEY_CLASSES_ROOT\*\shell\ConvertImage" /v "Icon" /f /d "D:\Documentos\Windows_Global_Scripts\convert-image\ico.ico"

@REM set "files=png jpg jpeg webp ico tiff"
@REM set "commands="
@REM for %%a in (%files%) do (
@REM   if "!commands!"=="" (
@REM     set "commands=%%a"
@REM   ) else (
@REM     set "commands=!commands!;%%a"
@REM   )
@REM   reg add "HKEY_CLASSES_ROOT\*\shell\ConvertImage\shell\%%a" /f /ve /d "%%a"
@REM   reg add "HKEY_CLASSES_ROOT\*\shell\ConvertImage\shell\%%a\command" /f /ve /d "\"D:\Documentos\Windows_Global_Scripts\convert-image\exe.exe\" \"%%1\" \"%%a\""
@REM )

@REM reg add "HKEY_CLASSES_ROOT\*\shell\ConvertImage" /v "SubCommands" /f /d "!commands!"


@REM echo comando criado no registro.
@REM pause