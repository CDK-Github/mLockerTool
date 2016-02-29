@echo off
SETLOCAL EnableDelayedExpansion
TITLE mLockerTool v1
ECHO Auteur : Cedric LEVEQUE
ECHO Bienvenue dans mLockerTool pour deverrouiller vos documents .doc
ECHO et vos documents .xls
ECHO __________________________________________________________________________
ECHO Configuration : 
Call :rouge "Attention ! Pensez a vider, au prealable, le repertoire de conversion et le repertoire final."
Call :rouge "Le mLockerTool.bat doit etre situe au meme endroit que le mLockerTool.vbs"
SET /p base= Repertoire de base (Ex : C:\chemin\base) :
SET /p trace= Repertoire de conversion (Ex : C:\chemin\trace) :
SET /p final= Repertoire final (Ex : C:\chemin\final) :
ECHO __________________________________________________________________________
ECHO Lancement et execution du programme...
FOR /f "delims=" %%a IN ('dir /b /s %base%') DO (
	start /wait "" cmd /c wscript mLockerTool.vbs %base% %trace% %final% "%%a"
)
ECHO Programme termine.
pause

:rouge
PowerShell -Command Write-Host %* -foreground "red" -background "black"
