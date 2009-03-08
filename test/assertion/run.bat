@echo off
dir /b test_*.vbs | cscript %~dp0..\..\bin\TestRunner.wsf //Job:ConsoleTestRunner /stdin+ %*
