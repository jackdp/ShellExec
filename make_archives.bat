@echo off
call globals.bat


copy %AppExe32Compiled% %AppExe%
%AppExe% --help > %README%
if exist %PortableFileZip32% del %PortableFileZip32%
%CreatePortableZip32%



copy %AppExe64Compiled% %AppExe%
%AppExe% --help > %README%
if exist %PortableFileZip64% del %PortableFileZip64%
%CreatePortableZip64%



copy %AppExe64Compiled% %AppExe%

