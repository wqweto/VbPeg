@echo off
setlocal
set root_dir=%~dp0
set runner_dir=%~dp0Runner
set vb_exe="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
set diff_exe="%ProgramFiles%\Git\usr\bin\diff.exe" --color=auto -u
set vbpeg_exe="%root_dir%..\vbpeg.exe" -q
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
    for /f "usebackq delims=" %%i in (`for /f %%k in ^('dir /b /s /on "%root_dir%%1\*.peg"'^) do @if [%%~xk]^=^=[.peg] echo %%k`) do call :make_test "%%i"
    shift /1
    goto :param_loop
)
if [%last_test%]==[] (
    for /f "usebackq delims=" %%i in (`for /f %%k in ^('dir /b /s /on "%root_dir%\*.peg"'^) do @if [%%~xk]^=^=[.peg] echo %%k`) do call :make_test "%%i"
)

echo %count_run% tests run, %count_fail% tests failed
goto :eof

:make_test
pushd "%~dp1"
set /a count_run=%count_run% + 1
set test_failed=0
call :echo_nameext "%cd%"

:: generate parser module
del /q /s "%~dpn1.out" >nul 2>&1
%vbpeg_exe% "%~1" -o "%runner_dir%\mdParser.bas" >"%~dpn1.out" 2>&1
if errorlevel 1 (
    type "%~dpn1.out"
    goto :make_failed
)
call :compare_filesize "%~dpn1.out" 0
if errorlevel 1 (
    if [%upd_expect%]==[1] copy /y "%~dpn1.out" "%~dpn1.expect" >nul
    if exist "%~dpn1.expect" (
        %diff_exe% "%~dpn1.expect" "%~dpn1.out"
        if errorlevel 1 set test_failed=1
    ) else (
        type "%~dpn1.out"
        set test_failed=1
    )
) else (
    del /q /s "%~dpn1.expect" >nul 2>&1
)

:: compile runner project
del /q /s "%~dp1Runner.exe" >nul 2>&1
echo %1 compile results >"%temp%\vb6.out"
start "" /wait %vb_exe% /make "%runner_dir%\Runner.vbp" /outdir "%~dp1" /out "%temp%\vb6.out"
if not exist "%~dp1Runner.exe" type "%temp%\vb6.out"& goto :make_failed

:: run tests from *.in and diff *.out vs *.expect
for /f %%j in ('dir /b *.in') do (
    call :run_test "%~dp1Runner.exe" "%%j"
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
del /q /s "%~dpn2.out" >nul 2>&1
%1 %2 >"%~dpn2.out"
if [%upd_expect%]==[1] copy /y "%~dpn2.out" "%~dpn2.expect" >nul
if exist "%~dpn2.expect" (
    %diff_exe% "%~dpn2.expect" "%~dpn2.out"
    exit /b !errorlevel!
)
type "%~dpn2.out"
exit /b 1

:echo_nameext
echo %~nx1
goto :eof

:compare_filesize
if %~z1 gtr %2 exit /b 1
goto :eof
