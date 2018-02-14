@echo off
cls

.paket\paket.exe restore
if errorlevel 1 (
  exit /b %errorlevel%
)

cd .\src\ExcelProvider\
dotnet restore
cd ..\..\tests\ExcelProvider.Tests\
dotnet restore
cd ..\..\

packages\build\FAKE\tools\FAKE.exe build.fsx %*