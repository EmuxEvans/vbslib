@echo off

set basedir=%~dp0
set pkgname=vbslib-0.0.0

cd "%basedir%"
svn export . %pkgname%
cscript .\bin\tool.wsf //Job:MakeZip %pkgname%.zip %pkgname%
rmdir /s /q %pkgname%

dir %pkgname%.zip
