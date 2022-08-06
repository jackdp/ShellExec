@echo off

set PATH=%JP_MyToolsDir%;%JP_ToolsDir%;%PATH%

rem --------------- common -----------------
set AppName=ShellExec
set AppVer=1.0
set AppDate=2022.08.06
set AppFullName=%AppName% %AppVer%
set AppName_=ShellExec
set AppExe=ShellExec.exe
set AppUrl=https://www.pazera-software.com/products/shellexec/
set README=README.txt

::set ArchiveSrc=%AppFullName%_Project.7z


rem ----------------- Windows 32 bit ---------------------
set AppExe32Compiled=ShellExec.32.exe
set PortableFileZip32=%AppName_%_win32.zip
set CreatePortableZip32=7z a -tzip -mx=9 %PortableFileZip32% %AppExe% %README%


rem ----------------- Windows 64 bit ---------------------
set AppExe64Compiled=ShellExec.64.exe
set PortableFileZip64=%AppName_%_win64.zip
set CreatePortableZip64=7z a -tzip -mx=9 %PortableFileZip64% %AppExe% %README%


