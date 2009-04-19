@echo off

set basedir=%~dp0
cd "%basedir%"

cscript bin\tool.wsf //Nologo //Job:ShowDateTime "set pkgname=vbslib-%%Y%%m%%d" >mkpkg_tmp.bat
echo svn export . %%pkgname%% >>mkpkg_tmp.bat
echo cscript bin\tool.wsf //Job:MakeZip %%pkgname%%.zip %%pkgname%% >>mkpkg_tmp.bat
echo rmdir /s /q %%pkgname%% >>mkpkg_tmp.bat
echo dir %%pkgname%%.zip >>mkpkg_tmp.bat

call mkpkg_tmp.bat
del /s mkpkg_tmp.bat >nul
