let
// -------------------------------------------------------
pqName = "fnTMSPriceListV1",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// -------------------------------------------------------
/*
Parametrar:
ServiceArea_p           -
MaxDate_p               -
MinDate_p               -
FilePath_p              -
DBQuery_p               -

*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-06-14    tswm00    Ny funktion
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),
        
    ERR_InvalidServiceArea = preERRORmsg & "Invalid ServiceArea",
    ERR_InvalidMaxDate = preERRORmsg & "Invalid maxdate",
    ERR_InvalidFilePath = preERRORmsg & "Invalid filepath",
    ERR_InvalidDBQuery = preERRORmsg & "Invalid DBQuery",

// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 4,
//    1: XX
//    2: xxx
//    3: xx


    ServiceArea_p = 
    //--------------------------------------
    if TestScenario = 1 then
        "IAM"
    else if TestScenario = 2 then
        "CLIENT"
    else if TestScenario = 3 then
        "NETWORK"
    else if TestScenario = 4 then
        "SYSTEM"
    else
        null,

    MaxDate_p = 
    //--------------------------------------
    if TestScenario = 1 then
        DateTime.From("2021-01-01")
    else if TestScenario = 2 then
        DateTime.From("2021-01-01")
    else if TestScenario = 3 then
        DateTime.From("2021-01-01")
    else if TestScenario = 4 then
        DateTime.From("2022-01-01")
    else
        null,

    MinDate_p =
    //-------------------------------------
    if TestScenario = 1 then
        DateTime.From("2021-01-01")
    else if TestScenario = 2 then
        DateTime.From("2021-01-01")
    else if TestScenario = 3 then
        DateTime.From("2021-01-01")
    else if TestScenario = 4 then
        DateTime.From("2022-01-01")
    else
        null,

    FilePath_p =
    //-------------------------------------
    if TestScenario = 1 then
        "\\ad.karlstad.se\VS\Filer\KLK\DebiteringssystemIT\Tjänsteområden\TM2017Common\TM2017Prislista.accdb"
    else if TestScenario = 2 then
        "\\ad.karlstad.se\VS\Filer\KLK\DebiteringssystemIT\Tjänsteområden\TM2017Common\TM2017Prislista.accdb"
    else if TestScenario = 3 then
        "\\ad.karlstad.se\VS\Filer\KLK\DebiteringssystemIT\Tjänsteområden\TM2017Common\TM2017Prislista.accdb"
    else if TestScenario = 4 then
        "\\ad.karlstad.se\VS\Filer\KLK\DebiteringssystemIT\Tjänsteområden\TM2017Common\TM2017Prislista.accdb"
    else
        null,

    DBQuery_p =
    //-------------------------------------
    if TestScenario = 1 then
        "qComponentPriceList-V2"
    else if TestScenario = 2 then
        "qComponentPriceList-V2"
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "qComponentPriceList-V2"
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
    fnTMSPricelist =
    (
        ServiceArea_p   as              text,
        MinDate_p       as              datetime,
        MaxDate_p       as  nullable    datetime,
        FilePath_p      as              text,
        DBQuery_p       as              text
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------
        ServiceArea = 
        if ServiceArea_p <> null then
            Text.Upper(ServiceArea_p)
        else
            error ERR_InvalidServiceArea,
            
        MinDate = MinDate_p,
        
        MaxDate = 
        if MaxDate_p <> null then
            MaxDate_p
        else
            error ERR_InvalidMaxDate,
            
        FilePath = 
        if FilePath_p <> null then
            FilePath_p
        else
            error ERR_InvalidFilePath,
            
        DBQuery = 
        if DBQuery_p <> null then
            DBQuery_p
        else
            error ERR_InvalidDBQuery,


        SourceTbl =
        let
            Source = Access.Database(File.Contents(FilePath), [CreateNavigationProperties=true]),
            ViewsOnly = Table.SelectRows(Source, each ([Kind] = "View")),
            Qry = ViewsOnly{[Schema="",Item=DBQuery]}[Data]
        in
            Qry,

        SelectServiceArea = Table.SelectRows(SourceTbl, each ([#"CA-Description"] = ServiceArea)),

        SelectActive =
        if Table.HasColumns(SelectServiceArea,"Inactive") then
            Table.SelectRows(SelectServiceArea, each ([Inactive] = false))
        else
            SelectServiceArea,
            
        FilterDatesOlderThan =
        if MinDate <> null then
            Table.SelectRows(SelectActive, each [#"StCMP-ChangeDate"] >= MinDate)
        else
            SelectActive,

        DateValue = Table.AddColumn(FilterDatesOlderThan, "DateValue", each [#"StCMP-ChangeDate"]),
        UsedDateValue = Table.AddColumn(DateValue, "UsedDateValue", each MaxDate),

        ChangedTypeDateValue = Table.TransformColumnTypes(UsedDateValue,{{"DateValue", type number}, {"UsedDateValue", type number}}),

        DiffDateValue = Table.AddColumn(ChangedTypeDateValue, "DiffDateValue", each [DateValue]-[UsedDateValue]),
        RemoveFutureDates = Table.SelectRows(DiffDateValue, each [DiffDateValue] <= 0),

        Grouped = Table.Group(RemoveFutureDates, {"StCM-Denomination"}, {{"AllData", each _, type table}, {"MaxDiffDateValue", each List.Max([DiffDateValue]), type number}}),
        SelectedMax = Table.AddColumn(Grouped, "Custom", (x) => Table.SelectRows(x[AllData], each [DiffDateValue] = x[MaxDiffDateValue])),
        ExpandedCustom = Table.ExpandTableColumn(SelectedMax, "Custom", {"CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo"}, {"CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo"}),
        ReplacedNULL = Table.ReplaceValue(ExpandedCustom,null,0,Replacer.ReplaceValue,{"StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime"}),
        MonthlyCost = Table.AddColumn(ReplacedNULL, "MonthlyCost", each [#"StCMP-CommonCostKey"]*[#"StCMP-CommonCostValue"]/[#"StCMP-CommonCostDepTime"]+[#"StCMP-TOCostKey"]*[#"StCMP-TOCostValue"]/[#"StCMP-TOCostDepTime"]+[#"StCMP-OneTimeCost"]/[#"StCMP-OneTimeCostDepTime"]+[#"StCMP-PurchaseCost"]/[#"StCMP-PurchaseCostDepTime"]+[#"StCMP-InstallationCost"]/[#"StCMP-InstallationCostDepTime"]+[#"StCMP-OperatingCost"]/[#"StCMP-OperatingCostDepTime"], type number),
        SelectColumns = Table.SelectColumns(MonthlyCost,{"StCM-Denomination", "CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "MonthlyCost", "StCMP-ChangeDate", "StCMP-ChangeInfo"}),

        UppercasedText = Table.TransformColumns(SelectColumns,{{"StCM-Denomination", Text.Upper, type text}, {"CA-Description", Text.Upper, type text}, {"CSA-Description", Text.Upper, type text}}),
        ChangedType = Table.TransformColumnTypes(UppercasedText,{{"StCM-Denomination", type text}, {"CA-Description", type text}, {"CSA-Description", type text}, {"StCMP-CommonCostKey", type number}, {"StCMP-CommonCostValue", type number}, {"StCMP-CommonCostDepTime", type number}, {"StCMP-TOCostKey", type number}, {"StCMP-TOCostValue", type number}, {"StCMP-TOCostDepTime", type number}, {"StCMP-OneTimeCost", type number}, {"StCMP-OneTimeCostDepTime", type number}, {"StCMP-PurchaseCost", type number}, {"StCMP-PurchaseCostDepTime", type number}, {"StCMP-InstallationCost", type number}, {"StCMP-InstallationCostDepTime", type number}, {"StCMP-OperatingCost", type number}, {"StCMP-OperatingCostDepTime", type number}, {"StCMP-ChangeDate", type date}, {"StCMP-ChangeInfo", type text}, {"MonthlyCost", type number}})

    in
        ChangedType
///*
in
    fnTMSPricelist
//*/