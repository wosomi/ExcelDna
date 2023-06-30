setlocal

set PackageVersion=1.7.0-ac1
set PackageReferenceVersion=1.7.0-ac1
set DllVersion=1.7.0.3

set MSBuildPath="%MSBUILD_PATH%"

set rootPath=%~dp0..\..

set propsfile=%rootPath%\Directory.Build.props
copy /Y Directory.Build.props %propsfile%
PowerShell "(Get-Content %propsfile%) -replace '_VERSION_', '%DllVersion%' | Set-Content %propsfile%"
@if errorlevel 1 goto end

set targetsfile=%rootPath%\Directory.Build.targets.local
copy /Y PushAll.template.cmd PushAll.cmd
PowerShell "(Get-Content PushAll.cmd) -replace '_VERSION_', '%PackageVersion%' | Set-Content PushAll.cmd
@if errorlevel 1 goto end

set targetsfile=%rootPath%\Directory.Build.targets.local
copy /Y Directory.Build.targets %targetsfile%
PowerShell "(Get-Content %targetsfile%) -replace '_VERSION_', '%PackageReferenceVersion%' | Set-Content %targetsfile%"
@if errorlevel 1 goto end

cd %rootPath%\ExcelDna\Build
call BuildPackages.bat %PackageVersion% %DllVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\Registration\Build
copy /Y %targetsfile% %rootPath%\Registration\Source\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\IntelliSense\Build
copy /Y %targetsfile% %rootPath%\IntelliSense\Source\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

cd %rootPath%\ExcelDnaDoc\Build
copy /Y %targetsfile% %rootPath%\ExcelDnaDoc\Directory.Build.targets
call BuildPackages.bat %PackageVersion% %MSBuildPath%
@if errorlevel 1 goto end

:end
