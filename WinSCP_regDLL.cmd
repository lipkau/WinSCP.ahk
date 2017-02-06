@echo off
cls
pushd "%~dp0bin\"

%WINDIR%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe WinSCPnet.dll /codebase WinSCPnet.dll /tlb
echo.
echo.
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe WinSCPnet.dll /codebase WinSCPnet.dll /tlb
echo.
echo.
pause