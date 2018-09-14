### Manual push of packages

You can push the packages if you have permissions, either automatically using ``build Release`` or manually

    .\Build BuildPackage
    set APIKEY=...
    ..\fsharp\.nuget\nuget.exe push bin\ExcelProvider.*.nupkg %APIKEY% -Source https://nuget.org 

    git tag 1.0.0
    git push https://github.com/fsprojects/ExcelProvider --tags
