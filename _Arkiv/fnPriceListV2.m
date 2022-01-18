let
// -------------------------------------------------------
pqName = "fnPriceListV2",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// -------------------------------------------------------
/*
Parametrar:
    ServiceArea_p               Ange ServiceArea (IAM, APPLICATION, CLIENT, NETWORK, SYSTEM, TELE). Om null så returneras alla
    PricelistDBFile_p           Access DB fil
    PricelistDBQuery_p          Access DB query (PriceList)
    ComponentMappingDBQuery_p   Access DB query (ComponentMapping)
    Date_p                      Prislistedatum, senaste priset närmast före eller lika med detta datum
    DateMin_p                   Inkludera datum som är större än eller lika med detta datum. Om null inget urval
    DateMax_p                   Inkludera datum som är mindre än eller lika med detta datum. Om null inget urval
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2020-11-12    tswm00    New PQ.
2020-11-26    tswm00    Lägger till möjlighet att göra Component Mapping
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),
        
    MSG_InvalidDBfile = preERRORmsg & "Invalid DB file",
    MSG_InvalidDBquery = preERRORmsg & "Invalid DB query",
    MSG_InvalidDate = preERRORmsg & "Invalid Date_p",

// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 4,
//    1: Alla
//    2: Nät
//    3: Alla inga datumbegränsningar
//    4: O365 med ComponentMapping


    ServiceArea_p = 
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        "NETWORK"
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "APPLICATION"
    else
        null,

    PricelistDBFile_p = 
    //--------------------------------------
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

    PricelistDBQuery_p =
    //-------------------------------------
    if TestScenario = 1 then
        "qComponentPriceList-V2"
    else if TestScenario = 2 then
        "qNetworkComponentPricelist-V3"
    else if TestScenario = 3 then
        "qComponentPriceList-V2"
    else if TestScenario = 4 then
        "qComponentPriceList-V2"
    else
        null,

    ComponentMappingDBQuery_p =
    //-------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        "qComponentMapping-V1"
    else
        null,

    Date_p =
    //--------------------------------------
    if TestScenario = 1 then
        Date.From("2020-12-31")
    else if TestScenario = 2 then
        Date.From("2020-12-31")
    else if TestScenario = 3 then
        Date.From("2020-12-31")
    else if TestScenario = 4 then
        Date.From(DateTime.LocalNow())
    else
        null,
        
    DateMin_p =
    //--------------------------------------
    if TestScenario = 1 then
        Date.From("2020-01-01")
    else if TestScenario = 2 then
        Date.From("2020-01-01")
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        null
    else
        null,
        
    DateMax_p =
    //--------------------------------------
    if TestScenario = 1 then
        Date.From("2020-12-31")
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        null
    else if TestScenario = 4 then
        null
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
    fnPriceList =
    (
        ServiceArea_p             as             text,
        PricelistDBFile_p         as             text,
        PricelistDBQuery_p        as             text,
        ComponentMappingDBQuery_p as nullable    text,
        Date_p                    as             datetime,
        DateMin_p                 as nullable    datetime,
        DateMax_p                 as nullable    datetime
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------

        ServiceArea = Text.Upper(ServiceArea_p),

        PricelistDBFile = 
        if PricelistDBFile_p = null then
            error MSG_InvalidDBfile
        else
            PricelistDBFile_p,

        PricelistDBQuery = 
        if PricelistDBQuery_p = null then
            error MSG_InvalidDBquery
        else
            PricelistDBQuery_p,

        ComponentMappingDBQuery = ComponentMappingDBQuery_p,

        Date = 
        if Date_p = null then
            MSG_InvalidDate
        else
            DateTime.From(Date_p),

        DateMin = DateTime.From(DateMin_p),
        DateMax = DateTime.From(DateMax_p),

        // ---------------------------------------------------
        // Read and manage data 
        // ---------------------------------------------------
        PriceListTbl = 
            let
                SourceFile = Access.Database(File.Contents(PricelistDBFile)),
                SourceData = SourceFile{[Schema="",Item=PricelistDBQuery]}[Data],
                ExcludeInactive =
                if Table.HasColumns(SourceData,"Inactive") then
                    Table.SelectRows(SourceData, each ([Inactive] = false))
                else
                    SourceData,
                UppercasedText = Table.TransformColumns(ExcludeInactive,{{"StCM-Denomination", Text.Upper, type text}, {"CA-Description", Text.Upper, type text}, {"CSA-Description", Text.Upper, type text}})
            in
                UppercasedText,

        ComponentMappingTbl =
        if ComponentMappingDBQuery <> null then
            let
                SourceFile = Access.Database(File.Contents(PricelistDBFile)),
                SourceData = SourceFile{[Schema="",Item=ComponentMappingDBQuery]}[Data],
                UppercasedText = Table.TransformColumns(SourceData,{{"SpCM-Denomination", Text.Upper, type text}, {"StCM-Denomination", Text.Upper, type text}, {"CA-Description", Text.Upper, type text}, {"CSA-Description", Text.Upper, type text}})
            in
                UppercasedText
        else
            null,

        // ---------------------------------------------------
        // Make selections 
        // ---------------------------------------------------
        SelectServiceArea = 
        if ServiceArea <> null then
            Table.SelectRows(PriceListTbl, each ([#"CA-Description"] = ServiceArea))
        else
            PriceListTbl,

        SelectMinDate =
        if DateMin <> null then
            Table.SelectRows(SelectServiceArea, each ([#"StCMP-ChangeDate"] >= DateMin))
        else
            SelectServiceArea,

        SelectMaxDate =
        if DateMax <> null then
            Table.SelectRows(SelectMinDate, each ([#"StCMP-ChangeDate"] <= DateMax))
        else
            SelectMinDate,

        // ---------------------------------------------------
        // Create pricelist 
        // ---------------------------------------------------
        DateValue = Table.AddColumn(SelectMaxDate, "DateValue", each [#"StCMP-ChangeDate"]),
        UsedDateValue = Table.AddColumn(DateValue, "UsedDateValue", each Date),
        ChangedTypeDateValue = Table.TransformColumnTypes(UsedDateValue,{{"DateValue", type number}, {"UsedDateValue", type number}}),

        DiffDateValue = Table.AddColumn(ChangedTypeDateValue, "DiffDateValue", each [DateValue]-[UsedDateValue]),
        RemoveFutureDates = Table.SelectRows(DiffDateValue, each [DiffDateValue] <= 0),

        Grouped = Table.Group(RemoveFutureDates, {"StCM-Denomination"}, {{"AllData", each _, type table}, {"MaxDiffDateValue", each List.Max([DiffDateValue]), type number}}),
        SelectedMax = Table.AddColumn(Grouped, "Custom", (x) => Table.SelectRows(x[AllData], each [DiffDateValue] = x[MaxDiffDateValue])),
        RemoveColumns1 = Table.RemoveColumns(SelectedMax,{"AllData", "MaxDiffDateValue"}),
        ExpandedCustom = Table.ExpandTableColumn(RemoveColumns1, "Custom", {"CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo", "DateValue", "UsedDateValue", "DiffDateValue"}, {"CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo", "DateValue", "UsedDateValue", "DiffDateValue"}),
        
        // ---------------------------------------------------
        // Map components (TMS)
        // ---------------------------------------------------
        DoMapComponents = ComponentMappingTbl <> null,
        MapComponents =
        if DoMapComponents then
            let
                MergeComponentMapping = Table.NestedJoin(ExpandedCustom, {"StCM-Denomination"}, ComponentMappingTbl, {"SpCM-Denomination"}, "ComponentMapping", JoinKind.LeftOuter),
                ExpandComponentMapping = Table.ExpandTableColumn(MergeComponentMapping, "ComponentMapping", {"StCM-Denomination"}, {"ComponentMapping.StCM-Denomination"}),
                StCMDenominationOLD = Table.RenameColumns(ExpandComponentMapping,{{"StCM-Denomination", "StCM-DenominationOLD"}}),
                ComponentMapped = Table.AddColumn(StCMDenominationOLD, "ComponentMapped", each [#"ComponentMapping.StCM-Denomination"] <> null, type logical),
                StCMDenomination = Table.AddColumn(ComponentMapped, "StCM-Denomination", each 
                if [ComponentMapped] then
                    [#"ComponentMapping.StCM-Denomination"]
                else
                    [#"StCM-DenominationOLD"])
            in
                StCMDenomination
        else
            Table.AddColumn(ExpandedCustom, "ComponentMapped", each false, type logical),
            
        // ---------------------------------------------------
        // Finish and result
        // ---------------------------------------------------
        OrganizeColumns = Table.SelectColumns(MapComponents,{"StCM-Denomination", "CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo"}),
        ChangedType = Table.TransformColumnTypes(OrganizeColumns,{{"StCM-Denomination", type text}, {"CA-Description", type text}, {"CSA-Description", type text}, {"StCMP-ChangeInfo", type text}, {"StCMP-CommonCostKey", type number}, {"StCMP-CommonCostValue", type number}, {"StCMP-TOCostKey", type number}, {"StCMP-TOCostValue", type number}, {"StCMP-OneTimeCost", type number}, {"StCMP-PurchaseCost", type number}, {"StCMP-InstallationCost", type number}, {"StCMP-OperatingCost", type number}, {"StCMP-CommonCostDepTime", Int64.Type}, {"StCMP-TOCostDepTime", Int64.Type}, {"StCMP-OneTimeCostDepTime", Int64.Type}, {"StCMP-PurchaseCostDepTime", Int64.Type}, {"StCMP-InstallationCostDepTime", Int64.Type}, {"StCMP-OperatingCostDepTime", Int64.Type}, {"StCMP-ChangeDate", type date}}),

        Result = ChangedType
    
    in
        Result
///*
in
    fnPriceList
//*/