let
    fnGetFileDate = (FullFilePath as text) as datetime => 
    let
        FilePath = fnReturnFilePath(FullFilePath),
        FileName = Text.End(FullFilePath,Text.Length(FullFilePath)-Text.Length(FilePath)),
        Source = Folder.Files(FilePath),
        #"Filtered Rows" = Table.SelectRows(Source, each ([Folder Path] = FilePath)),
        FileList = Table.RemoveColumns(#"Filtered Rows",{"Content", "Attributes"}),
        Date = DateTime.From(pqVLOOKUP(FileName,FileList,4,false))
    in
        Date
in
    fnGetFileDate