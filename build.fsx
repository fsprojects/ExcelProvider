#r "paket: groupref Build //"
#load ".fake/build.fsx/intellisense.fsx"

#if !Fake
#r "netstandard"
#endif

open Fake.Core
open Fake.Core.TargetOperators
open Fake.DotNet
open Fake.IO
open Fake.IO.FileSystemOperators
open Fake.IO.Globbing.Operators

//open Fake.Git
//open Fake.Testing.NUnit3
//open Fake.ReleaseNotesHelper
//open Fake.UserInputHelper
open System
open System.IO

Target.initEnvironment ()



// --------------------------------------------------------------------------------------
// START TODO: Provide project-specific details below
// --------------------------------------------------------------------------------------

// Information about the project are used
//  - for version and project name in generated AssemblyInfo file
//  - by the generated NuGet package
//  - to run tests and to publish documentation on GitHub gh-pages
//  - for documentation, you also need to edit info in "docs/tools/generate.fsx"

// The name of the project
// (used by attributes in AssemblyInfo, name of a NuGet package and directory in 'src')
let project = "ExcelProvider"

// Short summary of the project
// (used as description in AssemblyInfo and as a short summary for NuGet package)
let summary =
    "This library implements a read-only Excel type provider for Net Standard 2.0."


//// Git configuration (used for publishing documentation in gh-pages branch)
//// The profile where the project is posted
//let gitOwner = "fsprojects"
//let gitHome = "https://github.com/" + gitOwner

//// The name of the project on GitHub
//let gitName = "ExcelProvider"

//// The url for the raw files hosted
//let gitRaw = environVarOrDefault "gitRaw" "https://raw.github.com/fsprojects"

Trace.log "--Installing DotNet Core SDK if necessary"
let install = lazy DotNet.install DotNet.Versions.FromGlobalJson
let getSdkPath () = install.Value
Trace.log $"Value of getSdkPath = %A{getSdkPath}"


//let exec p args =
//    printfn "Executing %s %s" p args
//    Shell.Exec(p, args) |> function 0 -> () | d -> failwithf "%s %s exited with error %d" p args d

// Read additional information from the release notes document
let release = ReleaseNotes.load "RELEASE_NOTES.md"

// Helper active pattern for project types
let (|Fsproj|) (projFileName: string) =
    match projFileName with
    | f when f.EndsWith("fsproj") -> Fsproj
    | _ -> failwith (sprintf "Project file %s not supported. Unknown project type." projFileName)

// Generate assembly info files with the right version & up-to-date information
Target.create "AssemblyInfo" (fun _ ->
    Trace.log "--Creating new assembly files with appropriate version number and info"

    let getAssemblyInfoAttributes projectName =
        [ AssemblyInfo.Title(projectName)
          AssemblyInfo.Product project
          AssemblyInfo.Description summary
          AssemblyInfo.Version release.AssemblyVersion
          AssemblyInfo.FileVersion release.AssemblyVersion ]

    let getProjectDetails (projectPath: string) =
        let projectName = Path.GetFileNameWithoutExtension(projectPath)

        let directoryName = Path.GetDirectoryName(projectPath)
        let assemblyInfoAttributes = getAssemblyInfoAttributes projectName
        (projectPath, projectName, directoryName, assemblyInfoAttributes)

    !! "src/**/*.??proj"
    |> Seq.map getProjectDetails
    |> Seq.iter (fun (projFileName, _, folderName, attributes) ->
        match projFileName with
        | Fsproj ->
            let fileName = folderName + @"/" + "AssemblyInfo.fs"
            AssemblyInfoFile.createFSharp fileName attributes))

// Copies binaries from default VS location to expected bin folder
// But keeps a subdirectory structure for each project in the
// src folder to support multiple project outputs
Target.create "CopyBinaries" (fun _ ->
    Trace.log "-- Copy binaries to desired location"

    !! "src/**/*.??proj" -- "src/**/*.shproj"
    |> Seq.map (fun f ->
        (let source = (Path.GetDirectoryName f) </> "bin/Release"

         let target = "bin" </> (Path.GetFileNameWithoutExtension f)

         source, target))
    |> Seq.iter (fun (fromDir, toDir) -> Shell.copyDir toDir fromDir (fun _ -> true)))


// --------------------------------------------------------------------------------------
// Clean build results
Target.create "Clean" (fun _ ->
    Trace.log "--Cleaning various directories"

    !! "bin"
    ++ "temp"
    ++ "tmp"
    ++ "test/bin"
    ++ "test/obj"
    ++ "src/**/bin"
    ++ "src/**/obj"
    |> Shell.cleanDirs)





//Target "CleanDocs" (fun _ ->
//    CleanDirs ["docs/output"]
//)

// --------------------------------------------------------------------------------------
// Build library & test project

Target.create "Build" (fun _ ->
    Trace.log "--Building the binary files for distribution and the test project"

    let setParams (p: DotNet.BuildOptions) =
        { p with
            Configuration = DotNet.BuildConfiguration.Release }

    DotNet.build setParams "ExcelProvider.sln")


// --------------------------------------------------------------------------------------
// Run the unit tests using test runner
Target.create "RunTests" (fun _ ->
    Trace.log "-- Run the unit tests using test runner"
    let testProj = "./tests/ExcelProvider.Tests/ExcelProvider.Tests.fsproj"

    let testOptions (defaults: DotNet.TestOptions) =
        { defaults with
            Configuration = DotNet.BuildConfiguration.Release }

    DotNet.test testOptions testProj)

// --------------------------------------------------------------------------------------
// Build a NuGet package
Target.create "Nuget" (fun _ ->
    Trace.log "--Create the Nuget package using Paket pack"
    let inline dotnetSimple arg = DotNet.Options.lift install.Value arg

    let exitCode =
        DotNet.exec dotnetSimple "paket" "pack --template ./nuget/paket.template ./bin"

    if not exitCode.OK then
        failwithf "Process exit code '%d' <> 0. " exitCode.ExitCode)

//Target "PublishNuget" (fun _ ->
//    Paket.Push(fun p ->
//        { p with
//            WorkingDir = "bin" })
//)


//// --------------------------------------------------------------------------------------
//// Generate the documentation


//let fakePath = "packages" </> "build" </> "FAKE" </> "tools" </> "FAKE.exe"
//let fakeStartInfo script workingDirectory args fsiargs environmentVars =
//    (fun (info: System.Diagnostics.ProcessStartInfo) ->
//        info.FileName <- System.IO.Path.GetFullPath fakePath
//        info.Arguments <- sprintf "%s --fsiargs -d:FAKE %s \"%s\"" args fsiargs script
//        info.WorkingDirectory <- workingDirectory
//        let setVar k v =
//            info.EnvironmentVariables.[k] <- v
//        for (k, v) in environmentVars do
//            setVar k v
//        setVar "MSBuild" msBuildExe
//        setVar "GIT" Git.CommandHelper.gitPath
//        setVar "FSI" fsiPath)

///// Run the given buildscript with FAKE.exe
//let executeFAKEWithOutput workingDirectory script fsiargs envArgs =
//    let exitCode =
//        ExecProcessWithLambdas
//            (fakeStartInfo script workingDirectory "" fsiargs envArgs)
//            TimeSpan.MaxValue false ignore ignore
//    System.Threading.Thread.Sleep 1000
//    exitCode

//// Documentation
//let buildDocumentationTarget fsiargs target =
//    trace (sprintf "Building documentation (%s), this could take some time, please wait..." target)
//    let exit = executeFAKEWithOutput "docs/tools" "generate.fsx" fsiargs ["target", target]
//    if exit <> 0 then
//        failwith "generating reference documentation failed"
//    ()

//Target "GenerateReferenceDocs" (fun _ ->
//    buildDocumentationTarget "-d:RELEASE -d:REFERENCE" "Default"
//)

//let generateHelp' fail debug =
//    let args =
//        if debug then "--define:HELP"
//        else "--define:RELEASE --define:HELP"
//    try
//        buildDocumentationTarget args "Default"
//        traceImportant "Help generated"
//    with
//    | e when not fail ->
//        traceImportant "generating help documentation failed"

//let generateHelp fail =
//    generateHelp' fail false

//Target "GenerateHelp" (fun _ ->
//    DeleteFile "docs/content/release-notes.md"
//    CopyFile "docs/content/" "RELEASE_NOTES.md"
//    Rename "docs/content/release-notes.md" "docs/content/RELEASE_NOTES.md"

//    DeleteFile "docs/content/license.md"
//    CopyFile "docs/content/" "LICENSE.txt"
//    Rename "docs/content/license.md" "docs/content/LICENSE.txt"

//    generateHelp true
//)

//Target "GenerateDocs" DoNothing

//let createIndexFsx lang =
//    let content = """(*** hide ***)
//// This block of code is omitted in the generated HTML documentation. Use
//// it to define helpers that you do not want to show in the documentation.
//#I "../../../bin"

//(**
//F# Project Scaffold ({0})
//=========================
//*)
//"""
//    let targetDir = "docs/content" </> lang
//    let targetFile = targetDir </> "index.fsx"
//    ensureDirectory targetDir
//    System.IO.File.WriteAllText(targetFile, System.String.Format(content, lang))

//Target "AddLangDocs" (fun _ ->
//    let args = System.Environment.GetCommandLineArgs()
//    if args.Length < 4 then
//        failwith "Language not specified."

//    args.[3..]
//    |> Seq.iter (fun lang ->
//        if lang.Length <> 2 && lang.Length <> 3 then
//            failwithf "Language must be 2 or 3 characters (ex. 'de', 'fr', 'ja', 'gsw', etc.): %s" lang

//        let templateFileName = "template.cshtml"
//        let templateDir = "docs/tools/templates"
//        let langTemplateDir = templateDir </> lang
//        let langTemplateFileName = langTemplateDir </> templateFileName

//        if System.IO.File.Exists(langTemplateFileName) then
//            failwithf "Documents for specified language '%s' have already been added." lang

//        ensureDirectory langTemplateDir
//        Copy langTemplateDir [ templateDir </> templateFileName ]

//        createIndexFsx lang)
//)

//// --------------------------------------------------------------------------------------
//// Release Scripts

//Target "ReleaseDocs" (fun _ ->
//    let tempDocsDir = "temp/gh-pages"
//    CleanDir tempDocsDir
//    Repository.cloneSingleBranch "" (gitHome + "/" + gitName + ".git") "gh-pages" tempDocsDir

//    CopyRecursive "docs/output" tempDocsDir true |> tracefn "%A"
//    StageAll tempDocsDir
//    Git.Commit.Commit tempDocsDir (sprintf "Update generated documentation for version %s" release.NugetVersion)
//    Branches.push tempDocsDir
//)

//#load "paket-files/build/fsharp/FAKE/modules/Octokit/Octokit.fsx"
//open Octokit

//Target "Release" (fun _ ->
//    let user =
//        match getBuildParam "github-user" with
//        | s when not (String.IsNullOrWhiteSpace s) -> s
//        | _ -> getUserInput "Username: "
//    let pw =
//        match getBuildParam "github-pw" with
//        | s when not (String.IsNullOrWhiteSpace s) -> s
//        | _ -> getUserPassword "Password: "
//    let remote =
//        Git.CommandHelper.getGitResult "" "remote -v"
//        |> Seq.filter (fun (s: string) -> s.EndsWith("(push)"))
//        |> Seq.tryFind (fun (s: string) -> s.Contains(gitOwner + "/" + gitName))
//        |> function None -> gitHome + "/" + gitName | Some (s: string) -> s.Split().[0]

//    StageAll ""
//    Git.Commit.Commit "" (sprintf "Bump version to %s" release.NugetVersion)
//    Branches.pushBranch "" remote (Information.getBranchName "")

//    Branches.tag "" release.NugetVersion
//    Branches.pushTag "" remote release.NugetVersion

//    // release on github
//    createClient user pw
//    |> createDraft gitOwner gitName release.NugetVersion (release.SemVer.PreRelease <> None) release.Notes
//    // TODO: |> uploadFile "PATH_TO_FILE"
//    |> releaseDraft
//    |> Async.RunSynchronously
//)


let sourceFiles =
    !! "src/**/*.fs" ++ "src/**/*.fsi" ++ "build.fsx"
    -- "src/**/obj/**/*.fs"
    -- "src/AssemblyInfo*.fs"

Target.create "Format" (fun _ ->
    let result =
        sourceFiles
        |> Seq.map (sprintf "\"%s\"")
        |> String.concat " "
        |> DotNet.exec id "fantomas"

    if not result.OK then
        printfn "Errors while formatting all files: %A" result.Messages)

Target.create "CheckFormat" (fun _ ->
    let result =
        sourceFiles
        |> Seq.map (sprintf "\"%s\"")
        |> String.concat " "
        |> sprintf "%s --check"
        |> DotNet.exec id "fantomas"

    if result.ExitCode = 0 then
        Trace.log "No files need formatting"
    elif result.ExitCode = 99 then
        failwith "Some files need formatting, run `dotnet fake build -t Format` to format them"
    else
        Trace.logf "Errors while formatting: %A" result.Errors
        failwith "Unknown errors while formatting")



Target.create "BuildPackage" ignore

//// --------------------------------------------------------------------------------------
//// Run all targets by default. Invoke 'build <Target>' to override

Target.create "All" ignore

"Clean"
==> "AssemblyInfo"
==> "CheckFormat"
==> "Build"
==> "RunTests"
==> "CopyBinaries"
//  ==> "GenerateReferenceDocs"
//  ==> "GenerateDocs"
==> "All"
//  =?> ("ReleaseDocs",isLocalBuild)

"All" ==> "NuGet" ==> "BuildPackage"

//"CleanDocs"
//  ==> "GenerateHelp"
//  ==> "GenerateReferenceDocs"
//  ==> "GenerateDocs"

//"ReleaseDocs"
//  ==> "Release"

//"BuildPackage"
//  ==> "PublishNuget"
//  ==> "Release"

Target.runOrDefaultWithArguments "All"
