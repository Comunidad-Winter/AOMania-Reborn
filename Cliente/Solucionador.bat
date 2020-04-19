@echo off
cls
echo        .oo .oPYo. o     o               o              .oPYo.    .oPYo.       .oo
echo       .P 8 8    8 8b   d8                              8  .o8    8  .o8      .P 8
echo      .P  8 8    8 8`b d'8 .oPYo. odYo. o8 .oPYo.       8 .P'8    8 .P'8     .P  8
echo     oPooo8 8    8 8 `o' 8 .oooo8 8' `8  8 .oooo8       8.d' 8    8.d' 8         8
echo    .P    8 8    8 8     8 8    8 8   8  8 8    8       8o'  8    8o'  8         8
echo   .P     8 `YooP' 8     8 `YooP8 8   8  8 `YooP8       `YooP' 88 `YooP'88       8
echo.
echo Registrador de Libreria(s) del AoMania ReBorn
echo.
Pause
TIMEOUT 3
CD /d "%~dp0\libs"
echo Registrando Captura.ocx
RegSvr32 Captura.ocx
echo Registrando Comctl32.ocx
RegSvr32 COMCTL32.OCX
echo Registrando Cswsk32.ocx
RegSvr32 Cswsk32.ocx
echo Registrando listadoservers.ocx
RegSvr32 listadoservers.ocx
echo Registrando Mscomctl.ocx
RegSvr32 Mscomctl.ocx
echo Registrando Msinet.ocx
RegSvr32 MSINET.OCX
echo Registrando Mswinsck.ocx
RegSvr32 MSWINSCK.OCX
echo Registrando Richt32.ocx
RegSvr32 Richtx32.ocx
echo Registrando vbalprogbar6.ocx
RegSvr32 vbalProgBar6.ocx
echo Registrando dx8vb.dll
RegSvr32 dx8vb.dll
Pause.
echo.
Echo Libreria registrada.
Pause.