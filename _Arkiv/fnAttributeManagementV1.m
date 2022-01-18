let
// -------------------------------------------------------
pqName = "fnAttributeManagementV1",
// -------------------------------------------------------
// -------------------------------------------------------
/* Beskrivning
    Funktion för att mappa om, typecasta, och uppercase en datatabell baserat på en attributtabell
    Action_p sätts till null för full transformation. Genom att ange TYPECAST, ORGANIZE, UPPERCASE
        så fås enbart den transformationen och bygger på att kolumnerna redan är mappade enligt attributtbellen

    På sikt kan utökade funktioner införas för att förfina datahanteringen mha kolumnen "Value" där
    <TOM>:     Samma värde som mappad kolumn
    "value":   Värde (sträng, number, logical etc)
    [kolnamn]: Värdet skall tas från "kolnamn"
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-01-13    tswm00    New PQ
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    MSG_InvalidAttributeTbl =    preERRORmsg & "Invalid attribute table",
    MSG_InvalidDataTbl =         preERRORmsg & "Invalid data table",
    MSG_InvalidDataCol =         preERRORmsg & "Invalid data col",
    MSG_InvalidAction =         preERRORmsg & "Invalid action",

    Value_Organize =            "ORG",
    Value_Typecast =            "TYP",
    Value_Uppercase =           "UPP",
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 2,
//    1: tTESTAttributeMappingTable
//    2: SNAPSHOT data, Specific action

    AttributeTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        "tTESTAttributeMappingTable"
    else if TestScenario = 2 then
        "tTESTAttributeMappingTable"
    else
        null,

    DataTbl_p =
    //--------------------------------------
    if TestScenario = 1 then
        qReadXMLTEST
    else if TestScenario = 2 then
        fnReadSnapshotDataV4("ASSETREPORT")
    else
        null,

    Context_p =
    //--------------------------------------
    if TestScenario = 1 then
        "qAssetReportINTERFACELOAD"
    else if TestScenario = 2 then
        "qAssetReportINTERFACELOAD"
    else
        null,

    Action_p =
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        "upper"
    else
        null,
        
//*/
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
        else if Text.Start(Text.Upper(Action_p),3) = Value_Organize then
            Value_Organize
        else if Text.Start(Text.Upper(Action_p),3) = Value_Typecast then
            Value_Typecast
        else if Text.Start(Text.Upper(Action_p),3) = Value_Uppercase then
            Value_Uppercase
        else
            error MSG_InvalidAction,

// -------------------------------------------------------
// Data Management
// -------------------------------------------------------

        // Attribute Mapping Table
        //-----------------------------------------
        AttributeMappingTable = 
        let
            Source = Excel.CurrentWorkbook(){[Name=AttributeTbl]}[Content],
            ChangedType = Table.TransformColumnTypes(Source,{{"Context", type text}, {"attributeID", type text}, {"Include", type logical}, {"attributeValue", type text}, {"DataType", type text}, {"Value", type text}, {"Description", type text}}),
            FilteredRows = Table.SelectRows(ChangedType, each ([Include] = true) and ([Context]="*" or Text.Contains([Context], Context))),
            FilteredAttributeID = Table.SelectRows(FilteredRows, each [attributeID] <> null)
        in
            FilteredAttributeID,

        // Included attributes
        //-----------------------------------------
        AttributesIncluded = AttributeMappingTable[attributeValue],

        // Attribute Mapping
        //-----------------------------------------
        AttributeMapping =
        let
            Source = AttributeMappingTable,
            OrganizeColumns = Table.SelectColumns(Source,{"attributeID", "attributeValue"})
        in
            OrganizeColumns,

        // Attribute Mapping Types
        //-----------------------------------------
        AttributeMappingTypes = 
        let
            Source = AttributeMappingTable,
            SelectColumns = Table.SelectColumns(Source,{"attributeValue", "DataType"}),
            Transform = Table.TransformColumns(SelectColumns, {{"DataType", each Expression.Evaluate(_, [Currency.Type=Currency.Type, Int64.Type=Int64.Type, Percentage.Type=Percentage.Type]) }}),
            TableToRows = Table.ToRows(Transform)
        in
            TableToRows,

        // Attribute Transform list {function, type)
        //-----------------------------------------
        AttributeTransformList =
        let
            Source = AttributeMappingTable,
            SelectRows = Table.SelectRows(Source, each [DataType] = "type text"),
            TextUpper = Table.AddColumn(SelectRows, "TextUpper", each "Text.Upper"),
            SelectColumns1 = Table.SelectColumns(TextUpper,{"attributeValue", "TextUpper", "DataType"}),
            Transform1 = Table.TransformColumns(SelectColumns1, {{"DataType", each Expression.Evaluate(_, [Currency.Type=Currency.Type, Int64.Type=Int64.Type, Percentage.Type=Percentage.Type]) }}),
            Transform2 = Table.TransformColumns(Transform1, {{"TextUpper", each Expression.Evaluate(_,#shared)}}),
            ListOfList =Table.ToRows(Transform2)
        in
            ListOfList,

        // ---------------------------------------------------
        // Manage 
        // ---------------------------------------------------

    

        // ---------------------------------------------------
        // Result
        // ---------------------------------------------------

        Source = DataTbl,
        
        Transform =
        if Action = null then
            let
                // Mappa om attribute id till namn enligt extern tabell (MappingTbl)
                Rename = Table.RenameColumns(Source,Table.ToRows(AttributeMapping), MissingField.Ignore),

                // Typecasta enligt extern tabell (MappingTbl)
                TypeCast = Table.TransformColumnTypes(Rename, AttributeMappingTypes),

                // Uppercase av alla med "type text"
                Uppercase = Table.TransformColumns(TypeCast, AttributeTransformList),

                // Selektera kolumner som skall finnas med
                OrganizeColumns = Table.SelectColumns(Uppercase,AttributesIncluded)
            in
                OrganizeColumns
            
        else if Action = Value_Typecast then
                // Typecasta enligt extern tabell (MappingTbl)
                Table.TransformColumnTypes(Source, AttributeMappingTypes)

        else if Action = Value_Uppercase then
                // Uppercase av alla med "type text"
                Table.TransformColumns(Source, AttributeTransformList)
        
        else if Action = Value_Organize then
                // Selektera kolumner som skall finnas med
                Table.SelectColumns(Source,AttributesIncluded)
        else
            error MSG_InvalidAction,
        


        Result = Transform
    in
        Result
///*
in
    fnFilterText
//*/