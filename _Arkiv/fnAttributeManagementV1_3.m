let
// -------------------------------------------------------
pqName = "fnAttributeManagementV1_3",
// -------------------------------------------------------
// -------------------------------------------------------
/* Beskrivning
    Funktion för att mappa om, typecasta, och uppercase en datatabell baserat på en attributtabell
    Action_p sätts till null för full transformation. Genom att ange TYPECAST, ORGANIZE, UPPERCASE
        så fås enbart den transformationen och bygger på att kolumnerna redan är mappade enligt attributtbellen
        Genom att ange LIST så skapas en matchning mellan Attributtabell och datatabell för att se vilka attribut som ev. inte är mappade

    Via attributtabellen (Value) går det också att tilldela värden till nya kolumner enligt:
    <TOM>:     Samma värde som mappad kolumn
    "value":   Värde (sträng, number, logical etc). value KAN men måste inte vara inneslutet i "
    [kolnamn]: Värdet skall tas från "kolnamn"
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-01-13    tswm00    V1: New PQ
2021-01-18    tswm00    V1_1: Inkluderar värdetilldelning via tabell antingen som "value" eller [kolnamn]
2021-01-25    tswm00    V1_2: Lägger till hantering av "null" => riktig null
2021-02-10    tswm00    V1_3: Lägger till Action: LIST som ger en lista på mappade Attribut/Fält för given Context
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    MSG_InvalidAttributeTbl =    preERRORmsg & "Invalid attribute table",
    MSG_InvalidDataTbl =         preERRORmsg & "Invalid data table",
    MSG_InvalidDataCol =         preERRORmsg & "Invalid data col",
    MSG_InvalidAction =          preERRORmsg & "Invalid action",

    INFO_AttributeMapped =       preINFOmsg & "Attribute mapped",
    INFO_AttributeNotApplicable= preINFOmsg & "Attribute excluded",
    ERR_AttributeNotMapped =     preERRORmsg & "Attribute NOT mapped",
    ERR_AttributeMapError =      preERRORmsg & "Attribute mapping error",

    Action_Organize =            "ORG",
    Action_Typecast =            "TYP",
    Action_Uppercase =           "UPP",
    Action_List =                "LIST",

// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 5,
//    1: TestData1: Normal med full mappning
//    2: TestData2:, Specific action
//    3: MetaExtended
//    4: TestData3
//    5: Action List


    AttributeTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        "tTESTAttributeMappingTable"
    else if TestScenario = 2 then
        "tTESTAttributeMappingTable"
    else if TestScenario = 3 then
        "tAttributeMappingTable"
    else if TestScenario = 4 then
        "tTESTAttributeMappingTable"
    else if TestScenario = 5 then
        "tAttributeMappingTable"
    else
        null,

    DataTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        Excel.CurrentWorkbook(){[Name="tTestData1"]}[Content]
    else if TestScenario = 2 then
        Excel.CurrentWorkbook(){[Name="tTestData2"]}[Content]
    else if TestScenario = 3 then
        qReadMetaExtendedFile
    else if TestScenario = 4 then
        Excel.CurrentWorkbook(){[Name="tTestData3"]}[Content]
    else if TestScenario = 5 then
        qReadAssetReportFile
    else
        null,

    Context_p =
    //--------------------------------------
    if TestScenario = 1 then
        "TestData1"
    else if TestScenario = 2 then
        "TestData2"
    else if TestScenario = 3 then
        "qReadMetaExtendedINTERFACE"
    else if TestScenario = 4 then
        "TestData3"
    else if TestScenario = 5 then
        "qAssetReportINTERFACELOAD"
    else
        null,

    Action_p =
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        "typecast"
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        "List"
    else
        null,
        
*/
// =======================================================================
XXX_EXTERNAL_REFERENCES_XXX = null,
// =======================================================================

// =======================================================================
XXX_QUERY_BODY_XXX = null,
// =======================================================================
///*
    fnFilterText =
    (
                    AttributeTbl_p  as              text,    // Tabell med attributeinformation enligt bestämt format
                    DataTbl_p       as              table,   // Indatatabell
                    Context_p       as              text,    // Context enligt Context kolumn, tex querynamn, ServiceArea
        optional    Action_p        as    nullable  text     // null eller "TYPECAST", "ORGANIZE", UPPERCASE
    ) 
    as table =>


    let
//*/
    // ---------------------------------------------------
    // Manage parameters 
    // ---------------------------------------------------
        AttributeTbl = 
        if AttributeTbl_p = null then
            error MSG_InvalidAttributeTbl
        else
            AttributeTbl_p,

        DataTbl = 
        if DataTbl_p = null then
            error MSG_InvalidDataTbl
        else
            DataTbl_p,

        Context =
        if Context_p = null then
            "*"
        else
            Context_p,

        Action =
        if Action_p = null then
            null
        else if Text.Start(Text.Upper(Action_p),3) = Action_Organize then
            Action_Organize
        else if Text.Start(Text.Upper(Action_p),3) = Action_Typecast then
            Action_Typecast
        else if Text.Start(Text.Upper(Action_p),3) = Action_Uppercase then
            Action_Uppercase
        else if Text.Start(Text.Upper(Action_p),4) = Action_List then
            Action_List
        else
            error MSG_InvalidAction,

        //-----------------------------------------
        // Attribute Mapping Table complete
        //-----------------------------------------
        AttributeMappingTableComplete = 
        let
            Source = Excel.CurrentWorkbook(){[Name=AttributeTbl]}[Content],
            ChangedType = Table.TransformColumnTypes(Source,{{"Context", type text}, {"attributeID", type text}, {"Include", type logical}, {"attributeValue", type text}, {"DataType", type text}, {"Value", type text}, {"Description", type text}}),
            FilteredRows = Table.SelectRows(ChangedType, each ([Context]="*") or (Text.Contains([Context], Context))),

            IsColValue = 
            let
                AddColumn = Table.AddColumn(FilteredRows, "IsColValue", each Text.StartsWith([Value],"["), type logical),
                ReplaceNull = Table.ReplaceValue(AddColumn,null,false,Replacer.ReplaceValue,{"IsColValue"})
            in
                ReplaceNull,

            IsFixedValue = 
            let
                AddColumn = Table.AddColumn(IsColValue, "IsFixedValue", each [Value] <> null and not Text.StartsWith([Value],"["), type logical),
                ReplaceNull = Table.ReplaceValue(AddColumn,null,false,Replacer.ReplaceValue,{"IsFixedValue"})
            in
                ReplaceNull
        in
            IsFixedValue,

        //-----------------------------------------
        // Attribute Mapping Table
        //-----------------------------------------
        AttributeMappingTable = 
        let
            Source = AttributeMappingTableComplete,
            FilteredRows = Table.SelectRows(Source, each ([Include] = true))
        in
            FilteredRows,

        //-----------------------------------------
        // Included attributes
        //-----------------------------------------
        AttributesIncluded = 
        let
            Source = AttributeMappingTable,
            AttributesIncluded = Source[attributeValue]
        in
            AttributesIncluded,

        //-----------------------------------------
        // Attribute Mapping
        //-----------------------------------------
        AttributeMapping =
        let
            Source = AttributeMappingTable,
            OrganizeColumns = Table.SelectColumns(Source,{"attributeID", "attributeValue"})
        in
            OrganizeColumns,

        //-----------------------------------------
        // Attribute List
        //-----------------------------------------
        AttributeList =
        let
            Source = AttributeMappingTableComplete,
            OrganizeColumns = Table.SelectRows(Table.SelectColumns(Source,{"attributeID", "attributeValue", "Include"}), each [attributeID] <> null)
        in
            OrganizeColumns,

        //-----------------------------------------
        // Attribute rename/map list
        //-----------------------------------------
        AttributeRenameList = 
        let
            Source = AttributeMappingTable,
            SelectRows = Table.SelectRows(Source, each [attributeID] <> null and [IsColValue] = false and [IsFixedValue] = false),
            OrganizeColumns = Table.SelectColumns(SelectRows,{"attributeID", "attributeValue"}),
            TableToRows = Table.ToRows(OrganizeColumns)
        in
            TableToRows,

        //-----------------------------------------
        // Attribute mapping types list
        //-----------------------------------------
        AttributeMappingTypesList = 
        let
            Source = AttributeMappingTable,
            FilteredAttributeID = 
            if Action = null then
                Table.SelectRows(Source, each [attributeValue] <> null)
            else
                Source,
            SelectColumns = Table.SelectColumns(FilteredAttributeID,{"attributeValue", "DataType"}),
            Transform = Table.TransformColumns(SelectColumns, {{"DataType", each Expression.Evaluate(_, [Currency.Type=Currency.Type, Int64.Type=Int64.Type, Percentage.Type=Percentage.Type]) }}),
            TableToRows = Table.ToRows(Transform)
        in
            TableToRows,

        //-----------------------------------------
        // Attribute Transform list {function, type}
        //-----------------------------------------
        AttributeTransformList =
        let
            Source = AttributeMappingTable,
            FilteredAttributeID = 
            if Action = null then
                Table.SelectRows(Source, each [attributeID] <> null)
            else
                Source,
            SelectRows = Table.SelectRows(FilteredAttributeID, each [DataType] = "type text"),
            TextUpper = Table.AddColumn(SelectRows, "TextUpper", each "Text.Upper"),
            SelectColumns1 = Table.SelectColumns(TextUpper,{"attributeValue", "TextUpper", "DataType"}),
            Transform1 = Table.TransformColumns(SelectColumns1, {{"DataType", each Expression.Evaluate(_, [Currency.Type=Currency.Type, Int64.Type=Int64.Type, Percentage.Type=Percentage.Type]) }}),
            Transform2 = Table.TransformColumns(Transform1, {{"TextUpper", each Expression.Evaluate(_,#shared)}}),
            ListOfList = Table.ToRows(Transform2)
        in
            ListOfList,

        //-----------------------------------------
        // Attribute columns values
        //-----------------------------------------
        AttributeColumnValues =
        let
            // Lägg kolumnnamnet i en egen kolumn
            ColName = Table.AddColumn(AttributeMappingTable, "ColName", each
            if [IsColValue] then
                 Text.Middle([Value],1,Text.Length([Value])-2)
            else
                null, type text),
            SelectRows = Table.SelectRows(ColName , each ([IsColValue] = true)),
            SelectColumns = Table.SelectColumns(SelectRows,{"attributeValue", "ColName"})
        in
            SelectColumns,

        //-----------------------------------------
        // Attribute fixed values
        //-----------------------------------------
        AttributeFixedValues = 
        let
            // Ta bort ev " i början, slutet eller både av värdet
            ValueTEMP = Table.DuplicateColumn(AttributeMappingTable, "Value", "ValueTEMP"),
            SplitLeft = Table.SplitColumn(ValueTEMP, "ValueTEMP", Splitter.SplitTextByEachDelimiter({""""}, QuoteStyle.None, false), {"Value.1", "Value.2"}),
            SplitRight = Table.SplitColumn(SplitLeft, "Value.2", Splitter.SplitTextByEachDelimiter({""""}, QuoteStyle.None, true), {"Value.2", "Value.3"}),
            // Lägg det fasta värdet i en egen kolumn
            FixedValue = Table.AddColumn(SplitRight, "FixedValue", each
            if [IsFixedValue] then
                if [Value.1] <> null and [Value.1] <> "" then
                    [Value.1]
                else if [Value.2] <> null and [Value.2] <> "" then
                    [Value.2]
                else null
            else null),
            FilteredRows = Table.SelectRows(FixedValue, each ([IsFixedValue] = true)),
            SelectColumns = Table.SelectColumns(FilteredRows,{"attributeValue", "FixedValue"}),
            SetTrueNull = Table.ReplaceValue(SelectColumns,"null",null,Replacer.ReplaceValue,{"FixedValue"})
        in
            SetTrueNull,

        //-----------------------------------------
        // Data attribute list
        //-----------------------------------------
        DataAttributeList =
        let
            Columns = Table.ColumnNames(DataTbl),
            ConvertedToTable = Table.FromList(Columns, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            RenamedColumns = Table.RenameColumns(ConvertedToTable,{{"Column1", "DataTblAttributes"}})
        in
            RenamedColumns,

    // ---------------------------------------------------
    // Result
    // ---------------------------------------------------

    Transform =
        if Action = null then
            // Fullständig transformation
            let
                // Mappa om attribute id till namn enligt AttributeRenameList
                //---------------------------------------------------------------------------------------------------------------
                Rename = Table.RenameColumns(DataTbl,AttributeRenameList, MissingField.Ignore),

                // Lägg till kolumner med fixed värde enligt AttributeFixedValues
                //---------------------------------------------------------------------------------------------------------------
                AddColumnsFixedValue = List.Accumulate
                (
                    {0..List.Count(AttributeFixedValues[attributeValue])-1},
                    Rename,
                    (state, current) => 
                        Table.AddColumn(state, AttributeFixedValues[attributeValue]{current}, each AttributeFixedValues[FixedValue]{current})
                ),

                // Lägg till kolumner med kolumnvärde value enligt AttributeColumnValues
                //---------------------------------------------------------------------------------------------------------------
                AddColumnsColValue = List.Accumulate
                (
                    {0..List.Count(AttributeColumnValues[attributeValue])-1},
                    AddColumnsFixedValue,
                    (state, current) => 
                        Table.AddColumn(state, AttributeColumnValues[attributeValue]{current}, each Record.Field(_, AttributeColumnValues[ColName]{current}))
                ),

                // Typecasta enligt AttributeMappingTypesList
                //---------------------------------------------------------------------------------------------------------------
                TypeCast = Table.TransformColumnTypes(AddColumnsColValue, AttributeMappingTypesList),

                // Uppercase av alla med "type text" enligt AttributeTransformList
                //---------------------------------------------------------------------------------------------------------------
                Uppercase = Table.TransformColumns(TypeCast, AttributeTransformList),

                // Selektera kolumner enligt AttributesIncluded
                //---------------------------------------------------------------------------------------------------------------
                OrganizeColumns = Table.SelectColumns(Uppercase,AttributesIncluded, MissingField.Ignore)
            in
                OrganizeColumns
            
        else if Action = Action_Typecast then
            // Typecasta enligt AttributeMappingTypesList
            Table.TransformColumnTypes(DataTbl, AttributeMappingTypesList)

        else if Action = Action_Uppercase then
            // Uppercase av alla med "type text" enligt AttributeTransformList
            Table.TransformColumns(DataTbl, AttributeTransformList)
        
        else if Action = Action_Organize then
            // Selektera kolumner enligt AttributesIncluded
            Table.SelectColumns(DataTbl,AttributesIncluded)

        else if Action = Action_List then
            let
                MergeAttributeList = Table.NestedJoin(DataAttributeList, {"DataTblAttributes"}, AttributeList, {"attributeID"}, "AttributeList", JoinKind.LeftOuter),
                ExpandAttributeList = Table.ExpandTableColumn(MergeAttributeList, "AttributeList", {"attributeID", "attributeValue", "Include"}, {"AttributeList.attributeID", "AttributeList.attributeValue", "AttributeList.Include"}),
                Status = Table.AddColumn(ExpandAttributeList, "Status", each 
                if [AttributeList.Include] = true and [AttributeList.attributeValue] <> null then
                    INFO_AttributeMapped
                else if [AttributeList.Include] = false then
                    INFO_AttributeNotApplicable
                else if ([AttributeList.Include] = true and [AttributeList.attributeValue] = null) or ([AttributeList.Include] = null) then
                    ERR_AttributeNotMapped
                else
                    ERR_AttributeMapError, type text)
            in
                Status
        else
            error MSG_InvalidAction,

        Result = Transform
    in
        Result
///*
in
    fnFilterText
//*/