@echo off

cmd /v

setlocal
set TEST_LIST=
for %%t in (test_*.vbs) do set TEST_LIST=!TEST_LIST! %%t

cscript %~dp0..\bin\TestRunner.wsf //Job:ConsoleTestRunner %TEST_LIST%
