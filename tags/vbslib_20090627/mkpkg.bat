@echo off
setlocal

set basedir=%~dp0
cd "%basedir%"

for /f usebackq %%i in (
  `cscript bin\tool.wsf //Nologo //Job:ShowDateTime vbslib-%%Y%%m%%d`) do set pkgname=%%i

svn export http://vbslib.googlecode.com/svn/trunk %pkgname%
cscript bin\tool.wsf //Nologo //Job:Zip %pkgname%.zip %pkgname%
rmdir /s /q %pkgname%
