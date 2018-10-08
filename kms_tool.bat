@echo off
setlocal enabledelayedexpansion
title KMS BY imazes
REM color 4f
mode con cols=45 lines=30
REM set rr="HKCU\Console\%%SystemRoot%%_system32_cmd.exe"
REM reg add %rr% /v "WindowPosition" /t REG_DWORD /d 0x00C8012C /f >nul

openfiles >nul 2>&1
if %errorlevel% NEQ 0 goto :UACPrompt
cd /d %~dp0
goto A

:UACPrompt
Echo CreateObject^("Shell.Application"^).ShellExecute WScript.Arguments^(0^),"%*","","runas",1 >"%temp%\elevating.vbs"
%systemroot%\system32\cscript.exe //nologo "%temp%\elevating.vbs" "%~dpnx0"
del /f /q "%temp%\elevating.vbs" >nul 2>&1
exit /b

:A
cls
echo ==================起始菜单==================
set sel=0
set kmserver=windows.kms.app
echo                1:激活Windows
echo                2:激活Office
echo                3:查看激活状态
echo                *:退出
set /p sel=输入选项:
if "%sel%" equ "1" (
goto w
) else if "%sel%" equ "2" (
goto o
) else if "%sel%" equ "3" (
goto che
) else ( exit )

:w
cls
echo ========激活Windows========
set /p kmserver=请输入kms服务器:
set input=0
cls
echo =========激活选项=========
echo 1:激活Windows7        Pro Vol
echo.
echo 2:激活Windows8        Pro Vol
echo.
echo 3:激活Windows8.1      Pro Vol
echo.
echo 4:激活Windows10       Pro Vol
echo.
echo 5:激活Windows10       Enterprise Vol
echo.
echo 6:激活Windows10       LTSB 2016
echo.
echo 7:激活WindowsServer   Datacenter 2016
echo.
echo 8:激活Windows10       LTSC 2019
echo.
echo 直接回车: 执行通用激活
echo.
echo =========激活选项=========
set /p input=输入选项:
cls
if "%input%" equ "1" (
set kkey=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
) else if "%input%" equ "2" (
set kkey=NG4HW-VH26C-733KW-K6F98-J8CK4
) else if "%input%" equ "3" (
set kkey=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
) else if "%input%" equ "4" (
set kkey=W269N-WFGWX-YVC9B-4J6C9-T83GX
) else if "%input%" equ "5" (
set kkey=NPPR9-FWDCX-D2C8J-H872K-2YT43
) else if "%input%" equ "6" (
set kkey=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
) else if "%input%" equ "7" (
set kkey=CB7KF-BWN84-R7R2Y-793K2-8XDDG
) else if "%input%" equ "8" (
set kkey=M7XTQ-FN8P6-TTKYV-9D4CC-J462D
) else (goto act)

cls
cscript //Nologo %windir%\system32\slmgr.vbs /upk
ping/n 2 /w 1000 99.99 >nul
cscript //Nologo %windir%\system32\slmgr.vbs /ckms
ping/n 2 /w 1000 99.99 >nul
cscript //Nologo %windir%\system32\slmgr.vbs /ipk %kkey%

:act
cls
echo ===激活Windows......
cscript //Nologo %windir%\system32\slmgr.vbs /skms %kmserver%
cscript //Nologo %windir%\system32\slmgr.vbs /ato
echo =====完成激活
pause>nul
goto A

:o
cls
echo ========激活Office========
set /p kmserver=请输入kms服务器:
set input=0
cls
echo 1:激活Office 2016 (v_16)
echo 2:激活Office 2013 (v_15)
set /p input=输入选项
cls
if "%input%" equ "1" (
for /f "tokens=1,2,* " %%i in ('REG QUERY HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path ^| find /i "Path"') do set "offdir=%%k"
) else if "%input%" equ "2" (
for /f "tokens=1,2,* " %%i in ('REG QUERY HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path ^| find /i "Path"') do set "offdir=%%k"
) else (goto A)
cd /d !offdir!
cscript ospp.vbs /sethst:%kmserver%
cscript ospp.vbs /act
echo =====完成激活
pause>nul
goto A

:che
cls
echo ========查看激活状态========
set input=0
echo 1:查看Windows
echo 2:查看Office2016 16
echo 3:查看Office2013 15
set /p input=输入选项
cls
if "%input%" equ "1" (
slmgr /dlv
)else if "%input%" equ "2" (
for /f "tokens=1,2,* " %%i in ('REG QUERY HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path ^| find /i "Path"') do set "offdir=%%k"
cd /d !offdir!
cscript ospp.vbs /dstatus
) else if "%input%" equ "3" (
for /f "tokens=1,2,* " %%i in ('REG QUERY HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path ^| find /i "Path"') do set "offdir=%%k"
cd /d !offdir!
cscript ospp.vbs /dstatus) else (
echo 输入无效
goto A
)
echo ===查询完成
pause>nul
goto A