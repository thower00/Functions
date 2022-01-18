let
// =================================================================================================================================
// fnGetFileDateV2
// Returns date of file (File system)
// ---------------------------------------------------------------------------------------------------------------------------------
//  V1      2018-xx-xx  tswm00  New function
//  V2      2018-10-19  tswm00  Removes dependencies to other functions
// =================================================================================================================================

    fnGetFileDate =
    (
                    FullFilePath        as text,
        optional    ReturnDateColName   as nullable text    //Select "Date modified", "Date created" or "Date accessed"
                                                            // If omitted default value is set (DefaultReturnColName)
    )
    as datetime => 
    
    let
    
// =================================================================================================================================
// Constants
// =================================================================================================================================
    DefaultReturnColName = "Date modified",
// *********************************************************************************************************************************
// Main query
// *********************************************************************************************************************************

        ReturnDateColN = 
        if ReturnDateColName = null then
            DefaultReturnColName
        else
            ReturnDateColName,

        // Extract filname and filefolder
        FileName = Text.AfterDelimiter(FullFilePath, "\", {0, RelativePosition.FromEnd}),
        FilePath = Text.Start(FullFilePath, Text.Length(FullFilePath)-Text.Length(FileName)),

        // Get list of files in folder
        FolderContents = Folder.Files(FilePath),

        // Select only file(s) matching FileName and FilePath
        SelectFileAndPath = Table.SelectRows(FolderContents, each ([Name] = FileName) and ([Folder Path] = FilePath)),

        // To avoid error - remove duplicates of FileName (should not happen)
        RemoveDuplicateFileName= Table.Distinct(SelectFileAndPath, {"Name"}),

        // Pich result from column: ReturnDateColN
        Result = Record.Field(RemoveDuplicateFileName{0} , ReturnDateColN)
    in
        Result
in
    fnGetFileDate