module FSharp.Interop.Excel.ExcelProvider.ProviderImplementation

open System
open System.Collections.Generic
open System.IO
open FSharp.Interop.Excel
open FSharp.Interop.Excel.ExcelProvider
open Microsoft.FSharp.Core.CompilerServices
open ProviderImplementation.ProvidedTypes

[<AutoOpen>]
module internal Helpers =

    // Active patterns & operators for parsing strings
    let (@?) (s:string) i = if i >= s.Length then None else Some s.[i]

    let inline satisfies predicate (charOption:option<char>) = 
        match charOption with 
        | Some c when predicate c -> charOption 
        | _ -> None

    let (|EOF|_|) = function 
        | Some _ -> None
        | _ -> Some ()

    let (|LetterDigit|_|) = satisfies Char.IsLetterOrDigit
    let (|Upper|_|) = satisfies Char.IsUpper
    let (|Lower|_|) = satisfies Char.IsLower

    /// Turns a string into a nice PascalCase identifier
    let niceName (set:System.Collections.Generic.HashSet<_>) (s: string) =
        if s = s.ToUpper() then s else
        // Starting to parse a new segment 
        let rec restart i = seq {
            match s @? i with 
            | EOF -> ()
            | LetterDigit _ & Upper _ -> yield! upperStart i (i + 1)
            | LetterDigit _ -> yield! consume i false (i + 1)
            | _ -> yield! restart (i + 1) }

        // Parsed first upper case letter, continue either all lower or all upper
        and upperStart from i = seq {
            match s @? i with 
            | Upper _ -> yield! consume from true (i + 1) 
            | Lower _ -> yield! consume from false (i + 1) 
            | _ -> yield! restart (i + 1) }

        // Consume are letters of the same kind (either all lower or all upper)
        and consume from takeUpper i = seq {
            match s @? i with
            | Lower _ when not takeUpper -> yield! consume from takeUpper (i + 1)
            | Upper _ when takeUpper -> yield! consume from takeUpper (i + 1)
            | _ -> 
                yield from, i
                yield! restart i }

        // Split string into segments and turn them to PascalCase
        let mutable name =
            seq { for i1, i2 in restart 0 do 
                    let sub = s.Substring(i1, i2 - i1) 
                    if Seq.forall Char.IsLetterOrDigit sub then
                        yield sub.[0].ToString().ToUpper() + sub.ToLower().Substring(1) }
            |> String.concat ""

        while set.Contains name do 
          let mutable lastLetterPos = String.length name - 1
          while Char.IsDigit name.[lastLetterPos] && lastLetterPos > 0 do
            lastLetterPos <- lastLetterPos - 1
          if lastLetterPos = name.Length - 1 then
            name <- name + "2"
          elif lastLetterPos = 0 then
            name <- (UInt64.Parse name + 1UL).ToString()
          else
            let number = name.Substring(lastLetterPos + 1)
            name <- name.Substring(0, lastLetterPos + 1) + (UInt64.Parse number + 1UL).ToString()
        set.Add name |> ignore
        name


    let failInvalidCast fromObj (fromType:Type) (toType:Type) columnName rowIndex filename sheetname =
        sprintf
            "ExcelProvider: Cannot cast '%A' a '%s' to '%s'.\nFile: '%s'. Sheet: '%s'\nColumn '%s'. Row %i."
            fromObj
            fromType.FullName
            toType.FullName
            filename
            sheetname
            columnName
            rowIndex
        |> InvalidCastException
        |> raise

    // Avoids "warning FS0025: Incomplete pattern matches on this expression"
    // when using: (fun [row] -> <@@ ... @@>)
    let singleItemOrFail func items = 
        match items with
        | [ item ] -> func item
        | _ -> failwith "Expected single item list."

    // Avoids "warning FS0025: Incomplete pattern matches on this expression"
    // when using: (fun [row] -> <@@ ... @@>)
    let twoItemsOrFail func items = 
        match items with
        | [ a; b ] -> func a b
        | _ -> failwith "Expected two item list."
    
    // Avoids "warning FS0025: Incomplete pattern matches on this expression"
    // when using: (fun [row] -> <@@ ... @@>)
    let threeItemsOrFail func items = 
        match items with
        | [ a; b; c ] -> func a b c
        | _ -> failwith "Expected two item list."

    // Avoids "warning FS0025: Incomplete pattern matches on this expression"
    // when using: (fun [] -> <@@ ... @@>)
    let emptyListOrFail func items = 
        match items with
        | [] -> func()
        | _ -> failwith "Expected empty list"


    // get the type, and implementation of a getter property based on a template value
    let propertyImplementation columnIndex columnName (value : obj) =
        match value with
        | :? float -> typeof<double>, (fun row -> <@@ (%%row: Row).TryGetNullableValue<double> columnIndex columnName @@>)
        | :? bool -> typeof<bool>, (fun row -> <@@ (%%row: Row).TryGetNullableValue<bool> columnIndex columnName @@>)
        | :? DateTime -> typeof<DateTime>, (fun row -> <@@ (%%row: Row).TryGetNullableValue<DateTime> columnIndex columnName @@>)
        | :? string -> typeof<string>, (fun row -> <@@ (%%row: Row).TryGetValue<string> columnIndex columnName @@>)
        | _ -> typeof<obj>, (fun row -> <@@ (%%row: Row).GetValue columnIndex @@>)
        |> fun (cellType, getter) -> cellType, singleItemOrFail getter

    // gets a list of column definition information for the columns in a view
    let getColumnDefinitions (data : View) hasheaders forcestring =
        let getCell = getCellValue data
        [for columnIndex in 0 .. data.ColumnMappings.Count - 1 do
            let rawColumnName = if hasheaders then getCell 0 columnIndex |> string else "Column" + (columnIndex+1).ToString()
            if not (String.IsNullOrWhiteSpace(rawColumnName)) then
                let processedColumnName = rawColumnName.Replace("\n", "\\n")
                let cellType, getter =
                    if forcestring then
                        let getter = (fun row ->
                                        <@@
                                            let value = (%%row: Row).GetValue columnIndex |> string
                                            if String.IsNullOrEmpty value then null
                                            else value
                                        @@>) |> singleItemOrFail
                        typedefof<string>, getter
                    else
                        let cellValue = getCell (if hasheaders then 1 else 0) columnIndex
                        propertyImplementation columnIndex processedColumnName cellValue
                yield (processedColumnName, (columnIndex, cellType, getter))]


    let rootNamespace = "FSharp.Interop.Excel"

[<TypeProvider>]
type public ExcelProvider(cfg:TypeProviderConfig) as this =
    inherit TypeProviderForNamespaces(cfg, addDefaultProbingLocation=true, assemblyReplacementMap=[("ExcelProvider.DesignTime", "ExcelProvider.Runtime")])

    static let dict = Dictionary<_, _>()
    let executingAssembly = System.Reflection.Assembly.GetExecutingAssembly()

    // Create the main provided type
    let excelFileProvidedType = ProvidedTypeDefinition(executingAssembly, rootNamespace, "ExcelFile", Some(typeof<ExcelFileInternal>))

    /// Given a function to format names (such as `niceCamelName` or `nicePascalName`)
    /// returns a name generator that never returns duplicate name (by appending an
    /// index to already used names)
    /// 
    /// This function is curried and should be used with partial function application:
    ///
    ///     let makeUnique = uniqueGenerator nicePascalName
    ///     let n1 = makeUnique "sample-name"
    ///     let n2 = makeUnique "sample-name"
    ///
    let uniqueGenerator niceName =
      let set = new HashSet<_>()
      fun name ->
        let mutable name = niceName name
        while set.Contains name do 
          let mutable lastLetterPos = String.length name - 1
          while Char.IsDigit name.[lastLetterPos] && lastLetterPos > 0 do
            lastLetterPos <- lastLetterPos - 1
          if lastLetterPos = name.Length - 1 then
            if name.Contains " " then
                name <- name + " 2"
            else
                name <- name + "2"
          elif lastLetterPos = 0 && name.Length = 1 then
            name <- (UInt64.Parse name + 1UL).ToString()
          else
            let number = name.Substring(lastLetterPos + 1)
            name <- name.Substring(0, lastLetterPos + 1) + (UInt64.Parse number + 1UL).ToString()
        set.Add name |> ignore
        name

    let buildTypes (typeName:string) (args:obj[]) =
        let filename = args.[0] :?> string
        let sheetname = args.[1] :?> string
        let range = args.[2] :?> string
        let hasheaders = args.[3] :?> bool
        let forcestring = args.[4] :?> bool

        // resolve the filename relative to the resolution folder
        let resolvedFilename = Path.Combine(cfg.ResolutionFolder, filename)

        let ProvidedTypeDefinitionExcelCall (filename, sheetname, range, hasheaders, forcestring)  =
            let gen = uniqueGenerator id
            let data = openWorkbookView resolvedFilename sheetname range

            // define a provided type for each row, erasing to a int -> obj
            let providedRowType = ProvidedTypeDefinition("Row", Some(typeof<Row>))

            // add one property per Excel field
            let columnProperties = getColumnDefinitions data hasheaders forcestring
            for (columnName, (columnIndex, propertyType, getter)) in columnProperties do

                let prop = ProvidedProperty(columnName |> gen, propertyType, getterCode = getter)
                // Add metadata defining the property's location in the referenced file
                prop.AddDefinitionLocation((if hasheaders then 1 else 0), columnIndex, filename)
                providedRowType.AddMember(prop)

            // define the provided type, erasing to an seq<int -> obj>
            let providedExcelFileType = ProvidedTypeDefinition(executingAssembly, rootNamespace, typeName, Some(typeof<ExcelFileInternal>))

            // add a parameterless constructor which loads the file that was used to define the schema
            providedExcelFileType.AddMember(ProvidedConstructor([], invokeCode = emptyListOrFail (fun () -> <@@ ExcelFileInternal(resolvedFilename, sheetname, range, hasheaders) @@>)))

            // add a constructor taking the filename to load
            providedExcelFileType.AddMember(ProvidedConstructor([ProvidedParameter("filename", typeof<string>)], invokeCode = singleItemOrFail (fun filename -> <@@ ExcelFileInternal(%%filename, sheetname, range, hasheaders) @@>)))

             // add a constructor taking the filename and sheetname to load
            providedExcelFileType.AddMember(ProvidedConstructor([ProvidedParameter("filename", typeof<string>); ProvidedParameter("sheetname", typeof<string>)], invokeCode = twoItemsOrFail (fun fileName sheetname -> <@@ ExcelFileInternal(%%fileName, %%sheetname, range, hasheaders) @@>)))
            
            // add a constructor taking the stream to load
            providedExcelFileType.AddMember(ProvidedConstructor([ProvidedParameter("stream", typeof<Stream>); ProvidedParameter("format", typeof<ExcelFormat>)], invokeCode = twoItemsOrFail (fun stream format -> <@@ ExcelFileInternal(%%stream, %%format, sheetname, range, hasheaders) @@>)))

             // add a constructor taking the stream and sheetname to load
            providedExcelFileType.AddMember(ProvidedConstructor([ProvidedParameter("stream", typeof<Stream>); ProvidedParameter("format", typeof<ExcelFormat>); ProvidedParameter("sheetname", typeof<string>)], invokeCode = threeItemsOrFail (fun stream format sheetname -> <@@ ExcelFileInternal(%%stream, %%format, %%sheetname, range, hasheaders) @@>)))

            // add a new, more strongly typed Data property (which uses the existing property at runtime)
            providedExcelFileType.AddMember(ProvidedProperty("Data", typedefof<seq<_>>.MakeGenericType(providedRowType), getterCode = singleItemOrFail (fun excFile -> <@@ (%%excFile:ExcelFileInternal).Data @@>)))

            // add the row type as a nested type
            providedExcelFileType.AddMember(providedRowType)

            providedExcelFileType

        let key = (filename, sheetname, range, hasheaders, forcestring)
        if dict.ContainsKey(key) then dict.[key]
        else
            let res = ProvidedTypeDefinitionExcelCall key
            dict.[key] <- res
            res
    
    let parameters = 
        [ ProvidedStaticParameter("FileName", typeof<string>) 
          ProvidedStaticParameter("SheetName", typeof<string>, parameterDefaultValue = "") 
          ProvidedStaticParameter("Range", typeof<string>, parameterDefaultValue = "")
          ProvidedStaticParameter("HasHeaders", typeof<bool>, parameterDefaultValue = true)
          ProvidedStaticParameter("ForceString", typeof<bool>, parameterDefaultValue = false) ]

    let helpText = 
        """<summary>Typed representation of data in an Excel file.</summary>
           <param name='FileName'>Location of the Excel file.</param>
           <param name='SheetName'>Name of sheet containing data. Defaults to first sheet.</param>
           <param name='Range'>Specification using `A1:D3` type addresses of one or more ranges. Defaults to use whole sheet.</param>
           <param name='HasHeaders'>Whether the range contains the names of the columns as its first line.</param>
           <param name='ForceString'>Specifies forcing data to be processed as strings. Defaults to `false`.</param>"""

    do excelFileProvidedType.AddXmlDoc helpText
    do excelFileProvidedType.DefineStaticParameters(parameters, buildTypes)

    // add the type to the namespace
    do this.AddNamespace(rootNamespace,[excelFileProvidedType])

[<TypeProviderAssembly>]
do ()
