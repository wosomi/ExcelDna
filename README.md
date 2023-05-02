# Excel-DNA - Free and easy .NET for Excel

[![Version](https://img.shields.io/nuget/vpre/ExcelDna.AddIn.svg)](https://www.nuget.org/packages/ExcelDna.AddIn)
[![Downloads](https://img.shields.io/nuget/dt/ExcelDna.AddIn.svg)](https://www.nuget.org/packages/ExcelDna.AddIn)
[![GitHub contributors](https://img.shields.io/github/contributors/Excel-DNA/ExcelDna.svg)](https://github.com/Excel-DNA/ExcelDna/graphs/contributors)
[![License](https://img.shields.io/github/license/Excel-DNA/ExcelDna.svg)](https://github.com/Excel-DNA/ExcelDna/blob/master/LICENSE.txt)
[![Stack Overflow](https://img.shields.io/badge/stack%20overflow-excel--dna-orange.svg)](http://stackoverflow.com/questions/tagged/excel-dna)

This repository contains the core Excel-DNA library.

Please refer to
* [Excel-DNA Website](http://excel-dna.net), 
* [Excel-DNA Wiki Pages (on GitHub)](https://github.com/Excel-DNA/ExcelDna/wiki), and 
* [Excel-DNA Documentation](https://excel-dna.net/docs/introduction)

for additional information.

Support is on the [Excel-DNA Google group](https://groups.google.com/forum/#!forum/exceldna).

## Arup

This is a fork that makes one change to the base Dna to fix an issue where COM plugins are failing for some users. This is due to Dna trying to write to the machine hive when it does not have permission to. Dna tries to do this to account for the possibility of Excel being run "As Administrator", however, at Arup we do not need to worry about this. Therefore, we turn off this machine write hive feature by making the `RegistrationUtil.CanWriteMachineHive()` function always return `false`.

### Building

To build see [here](https://excel-dna.net/docs/guides-advanced/building-excedna-from-source).

NOTE: Before building making sure that:
1. You have installed the C++ modules on Visual Studio
2. You change the first line in the instructions, `set MSBuildPath="C:\Program Files\Microsoft Visual Studio\2022\Community\Msbuild\Current\Bin\amd64\MSBuild.exe"` to `set MSBuildPath="C:\Program Files\Microsoft Visual Studio\2022\Professional\Msbuild\Current\Bin\amd64\MSBuild.exe"` if using Visual Studio Professional
3. You have pulled down the non-Arup-Group versions of IntelliSense, Registration, and ExcelDnaDoc. There are not separate Arup-Group versions, but they are required to use this version locally
4. Sometimes having alternative nuget package sources can mess up the build script, turning them off (in Visual Studio Nuget Package Manager) seems to solve this issue
5. Make sure you are cd'd into the `MasterBuild` folder before running the script (if running from powershell/command line)
