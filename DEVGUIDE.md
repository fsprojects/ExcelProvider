### Manual push of packages

You can push the packages if you have permissions, either automatically using ``build Release`` or manually

    .\Build BuildPackage
    set APIKEY=...
    ..\fsharp\.nuget\nuget.exe push bin\ExcelProvider.0.9.1.nupkg %APIKEY% -Source https://nuget.org 

    git tag 0.9.1
    git push https://github.com/fsprojects/ExcelProvider --tags
