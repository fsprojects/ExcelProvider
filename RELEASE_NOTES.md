#### 0.3.3 - 09.01.2016
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