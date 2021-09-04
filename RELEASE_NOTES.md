#### 2.0.0 

* See the changes for the various release candidates below.


#### 2.0.0-rc3

* Add ability to read from stream instead of file

#### 2.0.0-rc2

* Add improved error messages when an invalid cast is attempted

#### 2.0.0-rc1

* Only build for NetStandard 2.0.  (reason for bump to V2)
* Convert paket to use dotnet tools infrastructure.
* Update FSharp.Formatting to version 11.4.2 and using dotnet tools (fsdocs command line).
* Update to dotnet 5.0 which is newest version of dotnet at time of writing.
* Update to FSharp.Core to 4.7.2.
* Fix issue when file locked by another process.

Still to be done in modernising the build infrastructure:
* Update the documentation generation and publishing.
* Update the publishing of nuget package to Nuget.
* Update the generation of releases on GitHub.

#### 1.0.1
* Fix bug in nuget package

#### 1.0.0
* Support .NET Core output
* Support .NET Standard output
* Support running in .NET SDK tools (.NET Core tools)
* Rename primary namespace to `FSharp.Interop.Excel`

#### 0.9.2
* Update ExcelDataReader

#### 0.9.1
* Correct dependencies

#### 0.9.0
* Upgrading ProvidedTypes.fs to latest FSharp.TypeProviders.SDK for  https://github.com/fsprojects/ExcelProvider/issues/52
* Give better release notes

#### 0.8.2 - 05.16.2017
* Upgrading ProvidedTypes.fs from FSharp.TypeProviders.SDK to fix a Mono5 compatibility issue.

#### 0.8.1 - 02.02.2017
* Clean up created temp folders by disposing ExcelDataReader - https://github.com/fsprojects/ExcelProvider/pull/37
* Upgrade to latest versions of dependencies
    ProvidedTypes.fs
    ProvidedTypes.fsi
    FAKE (4.50)
    Octokit (0.24)

#### 0.8.0 - 14.06.2016
* Support Row.GetValue("column") and handle of out-of-range column indexes explicitly - https://github.com/fsprojects/ExcelProvider/pull/27

#### 0.7.0 - 20.01.2016
* HasHeaders type provider parameter  - https://github.com/fsprojects/ExcelProvider/pull/26

#### 0.6.0 - 19.01.2016
* BREAKING CHANGE: Update handling of parameters - https://github.com/fsprojects/ExcelProvider/pull/25

#### 0.5.0 - 09.01.2016
* Using ProjectScaffold infrastructure
* Allow range parameter to take a sheet name - https://github.com/fsprojects/ExcelProvider/pull/18
* BUGFIX: Case insensitive ends with on EndsWith - https://github.com/fsprojects/ExcelProvider/pull/21
* Added unique property name generator for duplicated columns - https://github.com/fsprojects/ExcelProvider/pull/13
* Added support to multiline column names - https://github.com/fsprojects/ExcelProvider/pull/20
* Fixing handling of blank cells. Blank cells are treated as the default value for the inferred column type.
* Upgrading to the latest version of the ExcelDataReader library (2.1.2.3)
* Include nuget package dependencies

#### 0.1.0 - 20.03.2014
* Upgrading to the latest version of the ExcelDataReader library.
* Handling loading higher version of ICSharpCode.SharpZipLib.
* Defaulting to the first sheet in the spreadsheet if a sheet is not specifically indicated.
* Using a Row type, with a ToString() method

#### 0.0.1-alpha - 21.02.2014
* Initial release of ExcelProvider
