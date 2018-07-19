@echo off
setlocal
set root_dir=%~dp0
set runner_dir=%~dp0Runner
set vb_exe="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
set diff_exe="%ProgramFiles%\Git\usr\bin\diff.exe" --color=auto -u
set vbpeg_exe="%root_dir%..\vbpeg.exe"
set count_run=0
set count_fail=0
set upd_expect=0
set last_test=

:param_loop
if [%1]==[/expect] (
    set upd_expect=1
    shift /1
    goto :param_loop
)
if not [%1]==[] (
    set last_test=%1
    for /f %%i in ('dir /b /s %root_dir%\%1\*.peg') do call :make_test %%i
    shift /1
    goto :param_loop
)
if [%last_test%]==[] (
    for /f %%i in ('dir /b /s %root_dir%*.peg') do call :make_test %%i
)

echo %count_run% tests run, %count_fail% tests failed
goto :eof

:make_test
pushd %~dp1
set /a count_run=%count_run% + 1
set test_failed=0

:: generate parser module
%vbpeg_exe% -q %1 -o %runner_dir%\mdParser.bas
if errorlevel 1 goto :make_failed

:: compile runner project
del /q /s %~dp1Runner.exe >nul 2>&1
echo %1 compile results > %temp%\vb6.out
start "" /wait %vb_exe% /make %runner_dir%\Runner.vbp /outdir %~dp1 /out %temp%\vb6.out
if not exist %~dp1Runner.exe type %temp%\vb6.out& goto :make_failed

:: run tests from *.in and diff *.out vs *.expect
for /f %%j in ('dir /b *.in') do (
    call :run_test %~dp1Runner.exe %%j
    if errorlevel 1 set test_failed=1
)
if [%test_failed%]==[1] goto :make_failed

popd
goto :eof

:make_failed
set /a count_fail=%count_fail% + 1
popd
goto :eof

:run_test
del /q /s %~dpn2.out >nul 2>&1
%1 %2 > %~dpn2.out
if [%upd_expect%]==[1] copy /y %~dpn2.out %~dpn2.expect >nul
%diff_exe% %~dpn2.expect %~dpn2.out
exit /b %errorlevel%
