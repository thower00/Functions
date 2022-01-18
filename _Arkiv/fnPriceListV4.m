let
// -------------------------------------------------------
pqName = "fnPriceListV4",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// -------------------------------------------------------
/*
Parametrar:
    IndataTbl_p                 Indatatabell med nyckel som skall översättas till TMS och returnera priser för dessa
    IndataKeyCol_p              Indatatabellens nyckelkolumn
    ServiceArea_p               Ange ServiceArea (IAM, APPLICATION, CLIENT, NETWORK, SYSTEM, TELE). Om null så returneras alla
    PricelistDBFile_p           Access DB fil
    PricelistDBQuery_p          Access DB query (PriceList)
    TMSMappingDBQuery_p         Access DB query (ComponentMapping)
    TMSReMappingDBQuery_p       Access DB query (ComponentMapping)
    Date_p                      Prislistedatum, senaste priset närmast före eller lika med detta datum
    DateMin_p                   Inkludera datum som är större än eller lika med detta datum. Om null inget urval
    DateMax_p                   Inkludera datum som är mindre än eller lika med detta datum. Om null inget urval

Allt bygger på att de TMS som skall plockas fram finns prislistan och är prissatta där.
Ev. ommappning (både TMSMapping och TMSReMappning) sker INNAN uppslag mot prislistan. 

TMSMapping sker från någon specifik benämning (tex licenstyp) => TMS som sedan söks i prislistan
TMSReMapping ger en ommappning av TMS från TMSMapping till nytt TMS begrepp som sedan söks i prislistan

UseCase1:
-----------------------------------------------------------------------------
Indatatabell med specifika begrepp att mappa från skickas med.
TMSMapping måste göras och TMSMappingDBQuery_p måste skickas med. Befintliga 1:1 mappningar som finns i databasen skapar inga problem
TMSReMapping KAN göras om TMSReMappingDBQuery_p skickas med.

UseCase2:
-----------------------------------------------------------------------------
Ingen indatatabell. Vill bara ha ut prislistan från databasen enligt angivna parametrar (Date, DateMin, DateMax)
Prislistan sätt som indata.
TMSMapping KAN (och bör) användas om TMSMappingDBQuery_p skickas med.
TMSReMapping KAN användas om TMSReMappingDBQuery_pskickas med.

*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2020-11-12    tswm00    New PQ.
2020-11-26    tswm00    Lägger till möjlighet att göra Component Mapping
2021-04-23    tswm00    V3: Hanterar null värden
2021-11-17    tswm00    V4: Skriver om i grunden. Förändrar mappning och ger möjlighet att skicka med en tabell med specifika begrepp som skall översättas till TMS
                            Möjliggör också att göra TMSReMapping i querien.
                            Om ingen tabell med ingående
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg =   fnGetParameter("preERRORmsg"),
    preINFOmsg =    fnGetParameter("preINFOmsg"),

    ERR_InvalidIndataCol =          preERRORmsg & "Ogiltig indata tabell kolumnnamn",
    ERR_InvalidDBfile =             preERRORmsg & "Ogiltig databasfil",
    ERR_InvalidDBquery =            preERRORmsg & "Ogiltig databasquery",
    ERR_InvalidDate =               preERRORmsg & "Ogiltligt datum (Date_p)",
    ERR_TMSNotFoundInPriceList =    preERRORmsg & "TMS saknas i prislistan",
    
// =======================================================================
XXX_TEST_XXX = null,
// =======================================================================
/*
TestScenario = 4,
//    1: Bara prislista utan indatatabell. Utan DateMin/DateMax
//    2: O365 data med TMSReMapping. Utan DateMin/DateMax
//    3: O365 data med TMSReMapping. Med DateMin/DateMax
//    4: NÄT. Bara prislista utan indatatabell. Med DateMin/DateMax
//    5: Allt utom nät. Bara prislista utan indatatabell. Med DateMin/DateMax

    DBFile = "\\ad.karlstad.se\VS\Filer\KLK\DebiteringssystemIT\Tjänsteområden\TM2017Common\TM2017Prislista.accdb",

    IndataTbl_p = 
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        Table.Buffer(qO365ServiceDataREAD)
    else if TestScenario = 3 then
        Table.Buffer(qO365ServiceDataREAD)
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        null
    else
        null,

    IndataKeyCol_p = 
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        "LicenseMappingKey"
    else if TestScenario = 3 then
        "LicenseMappingKey"
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        null
    else
        null,

    ServiceArea_p = 
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        "APPLICATION"
    else if TestScenario = 3 then
        "APPLICATION"
    else if TestScenario = 4 then
        "NETWORK"
    else if TestScenario = 5 then
        null
    else
        null,

    PricelistDBFile_p = 
    //--------------------------------------
    if TestScenario = 1 then
        DBFile    
    else if TestScenario = 2 then
        DBFile    
    else if TestScenario = 3 then
        DBFile    
    else if TestScenario = 4 then
        DBFile    
    else if TestScenario = 5 then
        DBFile    
    else
        null,

    PricelistDBQuery_p =
    //-------------------------------------
    if TestScenario = 1 then
        "qComponentPriceList-V2"
    else if TestScenario = 2 then
        "qComponentPriceList-V2"
    else if TestScenario = 3 then
        "qComponentPriceList-V2"
    else if TestScenario = 4 then
        "qNetworkComponentPricelist-V3"
    else if TestScenario = 5 then
        "qComponentPriceList-V2"
    else
        null,

    TMSMappingDBQuery_p =
    //-------------------------------------
    if TestScenario = 1 then
        "qComponentMapping-V3"
    else if TestScenario = 2 then
        "qComponentMapping-V3"
    else if TestScenario = 3 then
        "qComponentMapping-V3"
    else if TestScenario = 4 then
        "qNetworkComponentMapping-V5"
    else if TestScenario = 5 then
        "qComponentMapping-V3"
    else
        null,

    TMSReMappingDBQuery_p =
    //-------------------------------------
    if TestScenario = 1 then
        "qServiceReMap-V2"
    else if TestScenario = 2 then
        "qServiceReMap-V2"
    else if TestScenario = 3 then
        "qServiceReMap-V2"
    else if TestScenario = 4 then
        "qServiceReMap-V2"
    else if TestScenario = 5 then
        "qServiceReMap-V2"
    else
        null,

    Date_p =
    //--------------------------------------
    if TestScenario = 1 then
        Date.From("2021-11-17")
    else if TestScenario = 2 then
        Date.From("2022-11-17")
    else if TestScenario = 3 then
        Date.From("2021-11-17")
    else if TestScenario = 4 then
        Date.From("2021-11-17")
    else if TestScenario = 5 then
        Date.From("2021-11-17")
    else
        null,
        
    DateMin_p =
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        Date.From("2021-01-01")
    else if TestScenario = 4 then
        Date.From("2021-01-01")
    else if TestScenario = 5 then
        Date.From("2021-01-01")
    else
        null,
        
    DateMax_p =
    //--------------------------------------
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        Date.From("2021-12-31")
    else if TestScenario = 4 then
        Date.From("2021-12-31")
    else if TestScenario = 5 then
        Date.From("2021-12-31")
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
        IndataTbl_p             as  nullable    table,
        IndataKeyCol_p          as  nullable    text,
        ServiceArea_p           as  nullable    text,
        PricelistDBFile_p       as              text,
        PricelistDBQuery_p      as              text,
        TMSMappingDBQuery_p     as  nullable    text,
        TMSReMappingDBQuery_p   as  nullable    text,
        Date_p                  as              datetime,
        DateMin_p               as  nullable    datetime,
        DateMax_p               as  nullable    datetime
    ) 
    as table =>

    let
//*/
        // ---------------------------------------------------
        // Manage parameters 
        // ---------------------------------------------------
        Is_IndataPriceList = IndataTbl_p = null,

        IndataTbl = 
        if Is_IndataPriceList then
        // Indatabellen är tom så sätt prislistan som indata
            IndataTblFromPriceList
        else
            IndataTbl_p,
        
        IndataKeyCol =
        if IndataKeyCol_p = null and Is_IndataPriceList then
        // Indata utgörs av prislistan. Sätt rätt kolumn
            "StCM-Denomination"
        else if IndataKeyCol_p <> null and not Is_IndataPriceList then
            IndataKeyCol_p

        else
            error ERR_InvalidIndataCol,

        ServiceArea = Text.Upper(ServiceArea_p),

        PricelistDBFile = 
        if PricelistDBFile_p = null then
            error ERR_InvalidDBfile
        else
            PricelistDBFile_p,

        PricelistDBQuery = 
        if PricelistDBQuery_p = null then
            error ERR_InvalidDBquery
        else
            PricelistDBQuery_p,

        TMSMappingDBQuery = TMSMappingDBQuery_p,
        Do_TMSMapping = TMSMappingDBQuery <> null,

        TMSReMappingDBQuery = TMSReMappingDBQuery_p,        
        Do_TMSReMapping = TMSReMappingDBQuery <> null,

        Date = 
        if Date_p = null then
            ERR_InvalidDate
        else
            DateTime.From(Date_p),

        DateMin = DateTime.From(DateMin_p),
        DateMax = DateTime.From(DateMax_p),

        // ---------------------------------------------------
        // Create indata table from PriceList 
        // ---------------------------------------------------
        IndataTblFromPriceList = 
        let
            Source = PriceListTbl,
            RemovedOtherColumns = Table.SelectColumns(Source,{"StCM-Denomination", "CA-Description", "CSA-Description"})
        in
            RemovedOtherColumns,

        // ---------------------------------------------------
        // Create Pricelist 
        // ---------------------------------------------------
        PriceListTbl = 
        let
            SourceFile = Access.Database(File.Contents(PricelistDBFile)),
            SourceData = SourceFile{[Schema="",Item=PricelistDBQuery]}[Data],
            ExcludeInactive =
            if Table.HasColumns(SourceData,"Inactive") then
            let
                ExcludeInactive = Table.SelectRows(SourceData, each ([Inactive] = false)),
                RemoveInactive = Table.RemoveColumns(ExcludeInactive,{"Inactive"})
            in
                RemoveInactive
            else
                SourceData,
            UppercasedText = Table.TransformColumns(ExcludeInactive,{{"StCM-Denomination", Text.Upper, type text}, {"CA-Description", Text.Upper, type text}, {"CSA-Description", Text.Upper, type text}}),
            ReplaceNULL1 = Table.ReplaceValue(UppercasedText,null,1,Replacer.ReplaceValue,{"StCMP-CommonCostDepTime", "StCMP-TOCostDepTime", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCostDepTime", "StCMP-OperatingCostDepTime"}),
            ReplaceNULL0 = Table.ReplaceValue(ReplaceNULL1,null,0,Replacer.ReplaceValue,{"StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-OneTimeCost", "StCMP-PurchaseCost", "StCMP-InstallationCost", "StCMP-OperatingCost"}),

            SelectMinDate = 
            if DateMin <> null then
                Table.SelectRows(ReplaceNULL0, each [#"StCMP-ChangeDate"] >= DateMin)
            else
                ReplaceNULL0,

            SelectMaxDate = 
            if DateMax <> null then
                Table.SelectRows(SelectMinDate, each [#"StCMP-ChangeDate"] <= DateMax)
            else
                SelectMinDate,

            DateValue = Table.AddColumn(SelectMaxDate, "DateValue", each [#"StCMP-ChangeDate"]),
            UsedDateValue = Table.AddColumn(DateValue, "UsedDateValue", each Date),
            ChangedTypeDateValue = Table.TransformColumnTypes(UsedDateValue,{{"DateValue", type number}, {"UsedDateValue", type number}}),

            DiffDateValue = Table.AddColumn(ChangedTypeDateValue, "DiffDateValue", each [DateValue]-[UsedDateValue]),
            RemoveFutureDates = Table.SelectRows(DiffDateValue, each [DiffDateValue] <= 0),

            Grouped = Table.Group(RemoveFutureDates, {"StCM-Denomination"}, {{"AllData", each _, type table}, {"MaxDiffDateValue", each List.Max([DiffDateValue]), type number}}),
            SelectedMax = Table.AddColumn(Grouped, "Custom", (x) => Table.SelectRows(x[AllData], each [DiffDateValue] = x[MaxDiffDateValue])),
            RemoveColumns1 = Table.RemoveColumns(SelectedMax,{"AllData", "MaxDiffDateValue"}),
            ExpandedCustom = Table.ExpandTableColumn(RemoveColumns1, "Custom", {"CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo", "DateValue", "UsedDateValue", "DiffDateValue"}, {"CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo", "DateValue", "UsedDateValue", "DiffDateValue"}),
            RemoveOtherColumns = Table.SelectColumns(ExpandedCustom,{"StCM-Denomination", "CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo"}),
            ChangeType = Table.TransformColumnTypes(RemoveOtherColumns,{{"StCM-Denomination", type text}, {"CA-Description", type text}, {"CSA-Description", type text}, {"StCMP-CommonCostKey", type number}, {"StCMP-CommonCostValue", type number}, {"StCMP-CommonCostDepTime", type number}, {"StCMP-TOCostKey", type number}, {"StCMP-TOCostValue", type number}, {"StCMP-TOCostDepTime", type number}, {"StCMP-OneTimeCost", type number}, {"StCMP-OneTimeCostDepTime", type number}, {"StCMP-PurchaseCost", type number}, {"StCMP-PurchaseCostDepTime", type number}, {"StCMP-InstallationCost", type number}, {"StCMP-InstallationCostDepTime", type number}, {"StCMP-OperatingCost", type number}, {"StCMP-OperatingCostDepTime", type number}, {"StCMP-ChangeDate", type datetime}, {"StCMP-ChangeInfo", type text}}),

            SelectServiceArea = 
            if ServiceArea <> null then
                Table.SelectRows(ChangeType, each ([#"CA-Description"] = ServiceArea))
            else
                ChangeType,

            SortedRows = Table.Sort(SelectServiceArea,{{"StCMP-ChangeDate", Order.Ascending}, {"StCM-Denomination", Order.Ascending}}),
            MonthlyCost = Table.AddColumn(SortedRows, "MonthlyCost", each [#"StCMP-CommonCostKey"]*[#"StCMP-CommonCostValue"]/[#"StCMP-CommonCostDepTime"]+[#"StCMP-TOCostKey"]*[#"StCMP-TOCostValue"]/[#"StCMP-TOCostDepTime"]+[#"StCMP-OneTimeCost"]/[#"StCMP-OneTimeCostDepTime"]+[#"StCMP-PurchaseCost"]/[#"StCMP-PurchaseCostDepTime"]+[#"StCMP-InstallationCost"]/[#"StCMP-InstallationCostDepTime"]+[#"StCMP-OperatingCost"]/[#"StCMP-OperatingCostDepTime"], type number),
            OrganizeColumns = Table.SelectColumns(MonthlyCost,{"StCM-Denomination", "CA-Description", "CSA-Description", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "StCMP-ChangeDate", "StCMP-ChangeInfo", "MonthlyCost"})
        in
            OrganizeColumns,

        // ---------------------------------------------------
        // Create TMSMappingTable 
        // ---------------------------------------------------
        TMSMappingTbl =
        let
            SourceFile = Access.Database(File.Contents(PricelistDBFile)),
            SourceData = SourceFile{[Schema="",Item=TMSMappingDBQuery]}[Data],
            UppercasedText = Table.TransformColumns(SourceData,{{"SpCM-Denomination", Text.Upper, type text}, {"StCM-Denomination", Text.Upper, type text}, {"CA-Description", Text.Upper, type text}, {"CSA-Description", Text.Upper, type text}}),

            SelectServiceArea = 
            if ServiceArea <> null then
                Table.SelectRows(UppercasedText, each ([#"CA-Description"] = ServiceArea))
            else
                UppercasedText,

            SelectWithinInterval = Table.SelectRows(SelectServiceArea, each (Date >= [ValidFromDate]) and (Date <= [ValidToDate]))
        in
            SelectWithinInterval,

        // ---------------------------------------------------
        // Create TMSReMappingTable 
        // ---------------------------------------------------
        TMSReMappingTbl =
        let
            SourceFile = Access.Database(File.Contents(PricelistDBFile)),
            SourceData = SourceFile{[Schema="",Item=TMSReMappingDBQuery]}[Data],
            UppercasedText = Table.TransformColumns(SourceData,{{"Domain", Text.Upper, type text}, {"Subdomain", Text.Upper, type text}, {"ReMapFromComponent", Text.Upper, type text}, {"ReMapToComponent", Text.Upper, type text}}),

            SelectServiceArea = 
            if ServiceArea <> null then
                Table.SelectRows(UppercasedText, each ([Domain] = ServiceArea))
            else
                UppercasedText,

            SelectWithinInterval = Table.SelectRows(SelectServiceArea, each (Date >= [ReMapFromDate]) and (Date <= [ReMapToDate]))
        in
            SelectWithinInterval,

        // ---------------------------------------------------
        // Start
        // ---------------------------------------------------
        DataSource = IndataTbl,
            
        // ---------------------------------------------------
        // Map TMS
        // ---------------------------------------------------
        MapTMS =
        if Do_TMSMapping then
        let
            MergeTMSMapping = Table.NestedJoin(DataSource, {IndataKeyCol}, TMSMappingTbl, {"SpCM-Denomination"}, "TMSMapping", JoinKind.LeftOuter),
            ExpandTMSMapping = Table.ExpandTableColumn(MergeTMSMapping, "TMSMapping", {"StCM-Denomination"}, {"fnPriceList.TMSMappedToTEMP"}),

            fnPriceListTMSMapped = Table.AddColumn(ExpandTMSMapping, "fnPriceList.TMSMapped", each [fnPriceList.TMSMappedToTEMP] <> null, type logical),
            fnPriceListTMSMappedTo = Table.AddColumn(fnPriceListTMSMapped, "fnPriceList.TMSMappedTo", each
            if [fnPriceList.TMSMapped] then
                [fnPriceList.TMSMappedToTEMP]
            else
                Record.Field(_,IndataKeyCol), type text),

            RemovefnPriceListTMSMappedTEMP = Table.RemoveColumns(fnPriceListTMSMappedTo,{"fnPriceList.TMSMappedToTEMP"})
        in
            RemovefnPriceListTMSMappedTEMP

        // Ingen TMS mapping skall göras. Lägg till kolumner med default värden
        else
        let
            fnPriceListTMSMapped = Table.AddColumn(DataSource, "fnPriceList.TMSMapped", each false, type logical),
            fnPriceListTMSMappedTo = Table.AddColumn(fnPriceListTMSMapped, "fnPriceList.TMSMappedTo", each [#"StCM-Denomination"], type text),
            RemoveOtherColumns = Table.SelectColumns(fnPriceListTMSMappedTo,{"StCM-Denomination", "CA-Description", "CSA-Description", "fnPriceList.TMSMapped", "fnPriceList.TMSMappedTo"})
        in
            RemoveOtherColumns,

        // ---------------------------------------------------
        // Remap TMS
        // ---------------------------------------------------
        ReMapTMS =
        if Do_TMSReMapping then
        let
            Source = MapTMS,
            MergeTMSReMapping = Table.NestedJoin(Source, {"fnPriceList.TMSMappedTo"}, TMSReMappingTbl, {"ReMapFromComponent"}, "TMSReMapping", JoinKind.LeftOuter),
            ExpandTMSReMapping = Table.ExpandTableColumn(MergeTMSReMapping, "TMSReMapping", {"ReMapToComponent"}, {"fnPriceList.TMSReMappedToTEMP"}),

            fnPriceListTMSReMapped = Table.AddColumn(ExpandTMSReMapping, "fnPriceList.TMSReMapped", each [fnPriceList.TMSReMappedToTEMP] <> null, type logical),
            fnPriceListTMSReMappedTo = Table.AddColumn(fnPriceListTMSReMapped, "fnPriceList.TMSReMappedTo", each
            if [fnPriceList.TMSReMapped] then
                [fnPriceList.TMSReMappedToTEMP]
            else
                [fnPriceList.TMSMappedTo], type text),

            RemovefnPriceListTMSReMappedTEMP = Table.RemoveColumns(fnPriceListTMSReMappedTo,{"fnPriceList.TMSReMappedToTEMP"})
        in
            RemovefnPriceListTMSReMappedTEMP

        // Ingen TMS Remapping skall göras. Lägg till kolumner med default värden
        else
        let
            Source = MapTMS,
            fnPriceListTMSReMapped = Table.AddColumn(Source, "fnPriceList.TMSReMapped", each false, type logical),
            fnPriceListTMSReMappedTo = Table.AddColumn(fnPriceListTMSReMapped, "fnPriceList.TMSReMappedTo", each [fnPriceList.TMSMappedTo], type text)
        in
            fnPriceListTMSReMappedTo,

        // ---------------------------------------------------
        // Look up in PriceList
        // ---------------------------------------------------
        MergePriceList = Table.NestedJoin(ReMapTMS, {"fnPriceList.TMSReMappedTo"}, PriceListTbl, {"StCM-Denomination"}, "PriceList", JoinKind.LeftOuter),
        ExpandPriceList = Table.ExpandTableColumn(MergePriceList, "PriceList", {"StCM-Denomination", "StCMP-CommonCostKey", "StCMP-CommonCostValue", "StCMP-CommonCostDepTime", "StCMP-TOCostKey", "StCMP-TOCostValue", "StCMP-TOCostDepTime", "StCMP-OneTimeCost", "StCMP-OneTimeCostDepTime", "StCMP-PurchaseCost", "StCMP-PurchaseCostDepTime", "StCMP-InstallationCost", "StCMP-InstallationCostDepTime", "StCMP-OperatingCost", "StCMP-OperatingCostDepTime", "MonthlyCost", "StCMP-ChangeDate", "StCMP-ChangeInfo"}, {"PriceList.StCM-Denomination", "PriceList.StCMP-CommonCostKey", "PriceList.StCMP-CommonCostValue", "PriceList.StCMP-CommonCostDepTime", "PriceList.StCMP-TOCostKey", "PriceList.StCMP-TOCostValue", "PriceList.StCMP-TOCostDepTime", "PriceList.StCMP-OneTimeCost", "PriceList.StCMP-OneTimeCostDepTime", "PriceList.StCMP-PurchaseCost", "PriceList.StCMP-PurchaseCostDepTime", "PriceList.StCMP-InstallationCost", "PriceList.StCMP-InstallationCostDepTime", "PriceList.StCMP-OperatingCost", "PriceList.StCMP-OperatingCostDepTime", "PriceList.MonthlyCost", "PriceList.StCMP-ChangeDate", "PriceList.StCMP-ChangeInfo"}),

        fnPriceListPriceListValid = Table.AddColumn(ExpandPriceList, "fnPriceList.PriceListValid", each [#"PriceList.StCM-Denomination"] <> null, type logical),
        NoTMSFoundInPriceList = Table.ReplaceValue(fnPriceListPriceListValid,null,ERR_TMSNotFoundInPriceList,Replacer.ReplaceValue,{"PriceList.StCM-Denomination"}),

        // ---------------------------------------------------
        // Finish and result
        // ---------------------------------------------------

        Result = NoTMSFoundInPriceList
    
    in
        Result
///*
in
    fnPriceList
//*/