@echo off
setlocal enabledelayedexpansion
set peg_dir=%~dp0peg
set runner_dir=%~dp0src
set runner_exe=%~dp0src\Runner.exe
set vb_exe="%ProgramFiles(x86)%\Microsoft Visual Studio\VB98\VB6.EXE"
set diff_exe="%ProgramFiles%\Git\usr\bin\diff.exe" --color=auto -u
set vbpeg_exe="%~dp0..\..\vbpeg.exe" -q
set replace_vbs=cscript //nologo "%runner_dir%\Replace.vbs"
set count_run=0
set count_fail=0
set upd_expect=
set test_modules=
set run_mode=

:main_param_loop
if [%1]==[expect] (
    set upd_expect=1
    shift /1
    goto :main_param_loop
)
if [%1]==[test] (
    echo Running tests...
    set run_mode=test_mode
    shift /1
    goto :main_param_loop
)
if [%run_mode%]==[] set run_mode=make_mode& echo Compiling parser classes...
goto :main_%run_mode%

:main_make_mode
if not [%1]==[] (
    for /f "usebackq delims=" %%i in (`for /f %%k in ^('dir /b /s /on "%peg_dir%\%1\*.peg"'^) do @if [%%~xk]^=^=[.peg] echo %%k`) do call :make_test "%%i"
    shift /1
    goto :main_param_loop
)
if "%test_modules%"=="" (
    for /f "usebackq delims=" %%i in (`for /f %%k in ^('dir /b /s /on "%peg_dir%\*.peg"'^) do @if [%%~xk]^=^=[.peg] echo %%k`) do call :make_test "%%i"
)
%replace_vbs% "/f:%runner_dir%\Runner.vbp.template" "/d:%runner_dir%\Runner.vbp" "/s:{test_modules}" "/r:%test_modules%"

echo Compiling runner.exe...
del /q /s "%runner_exe%" >nul 2>&1
echo %1 compile results >"%temp%\vb6.out"
start "" /wait %vb_exe% /make "%runner_dir%\Runner.vbp" /outdir "%runner_dir%" /out "%temp%\vb6.out"
if not exist "%runner_exe%" type "%temp%\vb6.out"& echo Compile failed
goto :eof

:main_test_mode
if not [%1]==[] (
    set test_modules=1
    for /f "usebackq delims=" %%i in (`for /f %%k in ^('dir /b /s /on "%peg_dir%\%1\*.peg"'^) do @if [%%~xk]^=^=[.peg] echo %%k`) do call :run_test "%%i"
    shift /1
    goto :main_param_loop
)
if [%test_modules%]==[] (
    for /f "usebackq delims=" %%i in (`for /f %%k in ^('dir /b /s /on "%peg_dir%\*.peg"'^) do @if [%%~xk]^=^=[.peg] echo %%k`) do call :run_test "%%i"
)
echo %count_run% tests run, %count_fail% tests failed
goto :eof

:make_test
pushd "%~dp1"
call :set_nameext "%cd%"
echo   %nameext%
set make_cls_file=%~dp1c%nameext%.cls
set make_out_file=%~dpn1.out
set make_expect_file=%~dpn1.expect

:: generate parser module
del /q /s "%make_out_file%" >nul 2>&1
%vbpeg_exe% "%~1" -o "%make_cls_file%" -set private=1 -set modulename=c%nameext% >"%make_out_file%" 2>&1
if errorlevel 1 (
    type "%make_out_file%"
    goto :make_failed
)
call :compare_filesize "%make_out_file%" 0
if errorlevel 1 (
    if [%upd_expect%]==[1] copy /y "%make_out_file%" "%make_expect_file%" >nul
    if exist "%make_expect_file%" (
        %diff_exe% "%make_expect_file%" "%make_out_file%"
    ) else (
        type "%make_out_file%"
    )
) else (
    del /q /s "%make_expect_file%" >nul 2>&1
)
if "%test_modules%"=="" (
    set test_modules=Class=c%nameext%; %make_cls_file%
) else (
    set test_modules=%test_modules:^=^^%^^pClass=c%nameext%; %make_cls_file%
)
:make_failed
popd
goto :eof

:run_test
pushd "%~dp1"
call :set_nameext "%cd%"
echo   %nameext%

:: run tests from *.in and diff *.out vs *.expect
for /f %%j in ('dir /b *.in') do (
    set /a count_run=!count_run! + 1
    call :diff_test "%runner_exe%" "c%nameext%" "%%j"
    if errorlevel 1 set /a count_fail=!count_fail! + 1
)
popd
goto :eof

:diff_test
set diff_out_file=%~dpn3.out
set diff_expect_file=%~dpn3.expect
del /q /s "%diff_out_file%" >nul 2>&1
%* >"%diff_out_file%" 2>&1
if errorlevel 1 (
    if [%upd_expect%]==[1] copy /y "%diff_out_file%" "%diff_expect_file%" >nul
    exit /b !errorlevel!
)
if [%upd_expect%]==[1] copy /y "%diff_out_file%" "%diff_expect_file%" >nul
if exist "%diff_expect_file%" (
    %diff_exe% "%diff_expect_file%" "%diff_out_file%"
    exit /b !errorlevel! 1
)
type "%diff_out_file%"
exit /b 1

:set_nameext
set nameext=%~nx1
goto :eof

:compare_filesize
if %~z1 gtr %2 exit /b 1
goto :eof
