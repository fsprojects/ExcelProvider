### Manual push of packages

You can push the packages if you have permissions, either automatically using ``build Release`` or manually

    set APIKEY=...
    ..\fsharp\.nuget\nuget.exe push bin\ExcelProvider.*.nupkg %APIKEY% -Source https://nuget.org 
