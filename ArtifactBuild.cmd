@echo off
pushd "%~dp0"
powershell Compress-7Zip "Macro\bin\Release" -ArchiveFileName "DcMacro.zip" -Format Zip
:exit
popd
@echo on
