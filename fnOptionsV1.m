let
// -------------------------------------------------------
pqName = "fnOptionsV1",
// -------------------------------------------------------
// Returns various result depending on Choice:
// -------------------------------------------------------
ChoiceTMSGROUPING     = "TMSGROUPING",        // Grupperar TMSName i TMSGroup.
// TableMerge         = kan anges. Om både TableMerge och ColumnMerge returneras tabellen med gruppvärden, annars bara grupperingstabellen
// ColumnMerge        = kan anges. Om både TableMerge och ColumnMerge returneras tabellen med gruppvärden, annars bara grupperingstabellen
// -------------------------------------------------------
ChoiceTMS             = "TMS",                // Matchar TMS tabell mot TMS options
// TableMerge         = anges
// ColumnMerge        = kan anges. Default enligt DEFAULT_ColumnMergeTMS
// -------------------------------------------------------
ChoiceCOST            = "COST",               // Matchar COST tabell mot COST options
// TableMerge         = anges
// ColumnMerge        = kan anges. Default enligt DEFAULT_ColumnMergeCOST
// -------------------------------------------------------
ChoiceTOALL           = "TO_ALL",             // Returnerar övergripande TO option (KEY, TECHNICAL COST, NONE för *)
// TableMerge         = anges ej
// ColumnMerge        = anges ej
// -------------------------------------------------------
ChoiceCOMMONALL       = "COMMON_ALL",         // Returnerar övergripande Common option (KEY, TECHNICAL COST, NONE för *)
// TableMerge         = anges ej
// ColumnMerge        = anges ej
// -------------------------------------------------------
/*
2020-04-06    tswm00    New PQ
2020-04-07    tswm00    Ändrar så att man skickar med en tabell och för tillbaks en smart-joinad tabell (TMS)
                        Gör samma för COST.
                        Lägger till två Choice: TO_ALL, COMMON_ALL
2020-04-08    tswm00    Utökar testhanteringen till flera testscenarios. Lägger till ColumnMerge med defaultvärde
2020-04-09    tswm00    Lägger till parametervald ColumnMerge även för Choice = "TMS"
2020-04-14    tswm00    Utökar COST tabellen med ExclCommon och ExclTO
2020-04-15    tswm00    Läser in också FixedTMSOperatingCost och FixedOperatingCostDepTime samt byter namn på utdata:
                        FixedTMSVolume => TMSVolume, FixedTMSOperatingCost => OperatingCost, FixedTMSOperatingCostDepTime => OperatingCostDepTime
2020-04-15    tswm00    Kompletterar med ExclTechnicalCost
2020-04-15    tswm00    Rättar till visa namn för TMS
2020-04-16    tswm00    Utökar så att det går att avsluta med wildcard (*) i slutet av TMSName, tex CONNECTION* för att sätta värden på alla som börjar med CONNECTION
2020-04-20    tswm00    Testar att lägga till TableMerge = Table.Buffer(TableMerge_p) för att se om det kan öka prestanda
2020-04-29    tswm00    Utökar så att val "TMSGROUPING" kan returnera tabell med gruppering eller bara grupptabellen
2020-04-29    tswm00    Rättar TMSGrouping så att jag kan hantera dynamiska gruppnamn
*/
// =======================================================================
XXX_CONSTANTS_XXX = null,
// =======================================================================
    preERRORmsg = fnGetParameter("preERRORmsg"),
    preINFOmsg = fnGetParameter("preINFOmsg"),

    ALLMarker = "DEFAULT",

    DEFAULT_ColumnMergeCOST = "COSTBeskrivning",
    DEFAULT_ColumnMergeTMS = "TMSName",
    DEFAULT_ExcludeBudget = false,
    DEFAULT_ExcludeTechnicalCost = false,
    DEFAULT_ExcludeCommon = false,
    DEFAULT_ExcludeTO = false,
        
// =======================================================================
XXX_TESTDATA_XXX = null,
// =======================================================================
/*
TestScenario = 9,
//    1: TMSGROUPING
//    2: TMS1
//    3: COST1
//    4: COST2
//    5: TOALL
//    6: COMMONALL
//    7: TMS2
//    8: TMS3 (qComponentPriceList)
//    9: TMSGROUPING (med tabell)
    
    Choice_p = 
    if TestScenario = 1 then
        ChoiceTMSGROUPING
    else if TestScenario = 2 then
        ChoiceTMS
    else if TestScenario = 3 then
        ChoiceCOST
    else if TestScenario = 4 then
        ChoiceCOST
    else if TestScenario = 5 then
        ChoiceTOALL
    else if TestScenario = 6 then
        ChoiceCOMMONALL
    else if TestScenario = 7 then
        ChoiceTMS
    else if TestScenario = 8 then
        ChoiceTMS
    else if TestScenario = 9 then
        ChoiceTMSGROUPING
    else
        null,

    TableMerge_p = 
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        TMSTestTable1
    else if TestScenario = 3 then
        COSTTestTable1
    else if TestScenario = 4 then
        COSTTestTable2
    else if TestScenario = 5 then
        null
    else if TestScenario = 6 then
        null
    else if TestScenario = 7 then
        TMSTestTable2
    else if TestScenario = 8 then
        TMSTestTable3
    else if TestScenario = 9 then
        TMSTestGroupingTable1
    else
        null,

    ColumnMerge_p =
    if TestScenario = 1 then
        null
    else if TestScenario = 2 then
        null
    else if TestScenario = 3 then
        "Beskrivning"
    else if TestScenario = 4 then
        null
    else if TestScenario = 5 then
        null
    else if TestScenario = 6 then
        null
    else if TestScenario = 7 then
        "TMSBeskrivning"
    else if TestScenario = 8 then
        "StCM-Denomination"
    else if TestScenario = 9 then
        "TMSType"
    else
        null,


    TMSTestGroupingTable1 = Table.FromRecords
    ({
    [TMSName = "TMS1",   TMSType = "CONNECTION",       Data1 = 0,    Data2 = "A"],
    [TMSName = "TMS2",   TMSType = "SWITCH",           Data1 = 1,    Data2 = "B"],
    [TMSName = "TMS3",   TMSType = "SWITCH",           Data1 = 2,    Data2 = "C"]
    }),


    TMSTestTable1 = Table.FromRecords
    ({
    [TMSName = "TILLÄGGSKOSTNAD:TILLÄGGSKOSTNAD",      Data1 = 0,    Data2 = "A"],
    [TMSName = "TMS1",                                 Data1 = 1,    Data2 = "B"],
    [TMSName = "TMS2",                                 Data1 = 2,    Data2 = "C"]
    }),

    TMSTestTable2 = Table.FromRecords
    ({
    [TMSBeskrivning = "TILLÄGGSKOSTNAD:TILLÄGGSKOSTNAD",      Data1 = 0,    Data2 = "A"],
    [TMSBeskrivning = "TMS1",                                 Data1 = 1,    Data2 = "B"],
    [TMSBeskrivning = "TMS2",                                 Data1 = 2,    Data2 = "C"]
    }),

    TMSTestTable3 = qComponentPriceList,

    COSTTestTable1 = Table.FromRecords
    ({
    [Beskrivning = "TILLÄGGSKOSTNAD:TILLÄGGSKOSTNAD",       Data1 = 0,    Data2 = "A"],
    [Beskrivning = "COST1",                                 Data1 = 1,    Data2 = "B"],
    [Beskrivning = "COST2",                                 Data1 = 2,    Data2 = "C"]
    }),

    COSTTestTable2 = fnTMSTMKCostV1("COST"),

*/
// =======================================================================
XXX_EXTERNAL_REFERENCES_XXX = null,
// =======================================================================
    fnStringIsInList = fnStringIsInListV2,

    TMSGroupingTblName = "tOptionsTMSGrouping",
    TMSGroupingTbl = 
        let
            Source = Excel.CurrentWorkbook(){[Name=TMSGroupingTblName]}[Content],
            ExcludeNULL = Table.SelectRows(Source, each ([TMSType] <> null)),
            ChangeType = Table.TransformColumnTypes(ExcludeNULL,{{"TMSType", type text}, {"TMSGroup1", type text}, {"TMSGroup2", type text}, {"TMSGroup3", type text}, {"TMSGroup4", type text}, {"TMSGroup5", type text}, {"TMSGroup6", type text}, {"TMSGroup7", type text}, {"TMSGroup8", type text}, {"TMSGroup9", type text}, {"TMSGroup10", type text}})
        in
            ChangeType ,

    TMSTblName = "tOptionsTMS",
    TMSTbl = 
        let
            Source = Excel.CurrentWorkbook(){[Name=TMSTblName]}[Content],
            ExcludeNULL = Table.SelectRows(Source, each ([TMSName] <> null)),
            TMSVolume =
            if Table.HasColumns(ExcludeNULL,"FixedTMS#(lf)Volume") then
                Table.RenameColumns(ExcludeNULL,{{"FixedTMS#(lf)Volume", "TMSVolume"}})
            else
                ExcludeNULL,

            OperatingCost = 
            if Table.HasColumns(TMSVolume,"FixedTMS#(lf)OperatingCost") then
                Table.RenameColumns(TMSVolume,{{"FixedTMS#(lf)OperatingCost", "OperatingCost"}})
            else
                TMSVolume,

            OperatingCostDepTime = 
            if Table.HasColumns(OperatingCost,"FixedTMS#(lf)OperatingCostDepTime") then
                Table.RenameColumns(OperatingCost,{{"FixedTMS#(lf)OperatingCostDepTime", "OperatingCostDepTime"}})
            else
                OperatingCost,

            DefaultExclTechnicalCost = Table.ReplaceValue(OperatingCostDepTime,null,DEFAULT_ExcludeTechnicalCost,Replacer.ReplaceValue,{"ExclTechnicalCost"}),
            ChangeType = Table.TransformColumnTypes(DefaultExclTechnicalCost,{{"TMSName", type text}, {"TOChoice", type text}, {"CommonChoice", type text}, {"ExclTechnicalCost", type logical}, {"TMSVolume", type number}, {"OperatingCost", type number}, {"OperatingCostDepTime", Int64.Type}})
        in
            ChangeType,

    COSTTblName = "tOptionsCOST",
    COSTTbl = 
        let
            Source = Excel.CurrentWorkbook(){[Name=COSTTblName]}[Content],
            ExcludeNULL = Table.SelectRows(Source, each ([COSTBeskrivning] <> null)),
            UpperCase = Table.TransformColumns(ExcludeNULL,{{"COSTBeskrivning", Text.Upper, type text}}),
            ChangeType = Table.TransformColumnTypes(UpperCase,{{"COSTBeskrivning", type text}, {"ExclBudget", type logical}, {"ExclTechnicalCost", type logical}, {"ExclCommon", type logical}, {"ExclTO", type logical}})
        in
            ChangeType,

// =======================================================================
XXX_QUERY_BODY_XXX = null,
// =======================================================================
///*
    fnOptions =
    (
        Choice_p        as             text,
        TableMerge_p    as nullable    table,
        ColumnMerge_p   as nullable    text
    ) 
    as any =>

    let
//*/
        Choice = Text.Upper(Choice_p),
        TableMerge = if TableMerge_p <> null then Table.Buffer(TableMerge_p) else TableMerge_p,
        ColumnMerge =
        if ColumnMerge_p = null then
            if Choice = ChoiceCOST then
                DEFAULT_ColumnMergeCOST
            else if Choice = ChoiceTMS then
                DEFAULT_ColumnMergeTMS
            else
                null
        else
            ColumnMerge_p,

    // -----------------------------------------------------------------
    TMSGrouping =
    // -----------------------------------------------------------------
        let
            Source = TMSGroupingTbl,

            KeptFirstRows = Table.FirstN(Source,1),
            RemovedOtherColumns1 = Table.SelectColumns(KeptFirstRows,{"TMSGroup1", "TMSGroup2", "TMSGroup3", "TMSGroup4", "TMSGroup5", "TMSGroup6", "TMSGroup7", "TMSGroup8", "TMSGroup9", "TMSGroup10"}),
            TransposedTable = Table.Transpose(RemovedOtherColumns1),
            ExcludeNULL = Table.SelectRows(TransposedTable, each ([Column1] <> null)),
            GroupNames = Table.ToList(ExcludeNULL),

            PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
            Unpivoted = Table.Unpivot(PromotedHeaders, GroupNames, "TMSTypeGroup", "Value"),
            RemovedOtherColumns = Table.SelectColumns(Unapivoted,{"TMSType", "TMSTypeGroup"}),
            ChangedType = Table.TransformColumnTypes(RemovedOtherColumns,{{"TMSType", type text}, {"TMSTypeGroup", type text}}),

            MergeGrouping =
            if TableMerge <> null or ColumnMerge <> null then
                let
                    Merge = Table.NestedJoin(TableMerge,{ColumnMerge},ChangedType,{"TMSType"},"Options",JoinKind.LeftOuter),
                    Expand = Table.ExpandTableColumn(Merge, "Options", {"TMSTypeGroup"}, {"Options.Group"})
                in
                    Expand
            else
                ChangedType
        in
            MergeGrouping,

    // -----------------------------------------------------------------
    TMS = 
    // -----------------------------------------------------------------
        let
            InitTable =
                let
                    Source = TMSTbl,
                    RemovedOtherColumns1 = Table.SelectColumns(Source,{"TMSName", "TOChoice", "CommonChoice", "ExclTechnicalCost", "TMSVolume", "OperatingCost", "OperatingCostDepTime"}),
                    ExcludeNULLTMSName = Table.SelectRows(RemovedOtherColumns1, each ([TMSName] <> null)),
                    ChangedType = Table.TransformColumnTypes(ExcludeNULLTMSName,{{"TMSName", type text}, {"TOChoice", type text}, {"CommonChoice", type text}, {"ExclTechnicalCost", type logical}, {"TMSVolume", type number}, {"OperatingCost", type number}, {"OperatingCostDepTime", Int64.Type}})
                in
                    ChangedType,

            // Lägg till ett antal kolumner som (ännu) inte finns i tabellen för ev. framtida bruk
            // ---------------------------------------------------------------------------------
            AddUnusedColumns =
                let
                    CommonKeyValue = if not Table.HasColumns(InitTable,"CommonKeyValue") then Table.AddColumn(InitTable, "CommonKeyValue", each null, type number) else InitTable,
                    TOKeyValue = if not Table.HasColumns(CommonKeyValue, "TOKeyValue") then Table.AddColumn(CommonKeyValue, "TOKeyValue", each null, type number) else CommonKeyValue,
                    CommonCostDepTime = if not Table.HasColumns(TOKeyValue, "CommonCostDepTime") then Table.AddColumn(TOKeyValue, "CommonCostDepTime", each null, type number) else TOKeyValue,
                    TOCostDepTime = if not Table.HasColumns(CommonCostDepTime, "TOCostDepTime") then Table.AddColumn(CommonCostDepTime, "TOCostDepTime", each null, type number) else CommonCostDepTime,
                    OneTimeCostDepTime = if not Table.HasColumns(TOCostDepTime, "OneTimeCostDepTime") then Table.AddColumn(TOCostDepTime, "OneTimeCostDepTime", each null, type number) else TOCostDepTime,
                    PurchaseCostDepTime = if not Table.HasColumns(OneTimeCostDepTime, "PurchaseCostDepTime") then Table.AddColumn(OneTimeCostDepTime, "PurchaseCostDepTime", each null, type number) else OneTimeCostDepTime,
                    InstallationCostDepTime = if not Table.HasColumns(PurchaseCostDepTime, "InstallationCostDepTime") then Table.AddColumn(PurchaseCostDepTime, "InstallationCostDepTime", each null, type number) else PurchaseCostDepTime
                in
                    InstallationCostDepTime,

            // Skapa den kompletta tabellen med radnummer för att sedan kunna merga ALLMarker raden från ALLTbl
            // ---------------------------------------------------------------------------------
            CompleteTbl = 
                let
                    OptionsTbl = AddUnusedColumns,

                    WildcardList =
                        let
                            DuplicatedColumn = Table.DuplicateColumn(OptionsTbl, "TMSName", "TMSName - Copy"),
                            SplitWildcard = Table.SplitColumn(DuplicatedColumn, "TMSName - Copy", Splitter.SplitTextByDelimiter("*", QuoteStyle.Csv), {"JoinString", "IsWildcard"}),
                            WildcardsOnly = Table.SelectRows(SplitWildcard, each ([IsWildcard] <> null)),
                            WildcardList = WildcardsOnly[JoinString]
                        in
                            WildcardList,
    
                    SourceColumn = ColumnMerge,
                    SourceTbl = TableMerge,

                    WildcardMatch = Table.AddColumn(SourceTbl, "WildcardMatch", each fnStringIsInList(Record.Field(_,SourceColumn),WildcardList, "VALUE"), type text),

                    NewSourceColumn = Table.AddColumn(WildcardMatch, "NewSourceColumn", each
                    if [WildcardMatch] = null then
                        Record.Field(_,SourceColumn)
                    else
                        [WildcardMatch],type text),

                    OptionsTblWithoutWildcard = Table.SplitColumn(OptionsTbl, "TMSName", Splitter.SplitTextByEachDelimiter({"*"}, QuoteStyle.Csv, true), {"TMSName"}),

                    MergeOptions = Table.NestedJoin(NewSourceColumn,{"NewSourceColumn"},OptionsTblWithoutWildcard,{"TMSName"},"Options",JoinKind.FullOuter),
                    ExpandOptions = Table.ExpandTableColumn(MergeOptions, "Options", {"TMSName", "TOChoice", "CommonChoice", "ExclTechnicalCost", "TMSVolume", "OperatingCost", "OperatingCostDepTime", "CommonKeyValue", "TOKeyValue", "CommonCostDepTime", "TOCostDepTime", "OneTimeCostDepTime", "PurchaseCostDepTime", "InstallationCostDepTime"}, {"Options.TMSName", "Options.TOChoice", "Options.CommonChoice", "Options.ExclTechnicalCost", "Options.TMSVolume", "Options.OperatingCost", "Options.OperatingCostDepTime", "Options.CommonKeyValue", "Options.TOKeyValue", "Options.CommonCostDepTime", "Options.TOCostDepTime", "Options.OneTimeCostDepTime", "Options.PurchaseCostDepTime", "Options.InstallationCostDepTime"}),
                    ColumnMergeOLD = Table.RenameColumns(ExpandOptions,{{ColumnMerge, ColumnMerge&"OLD"}}),

                    ColumnMergeNEW = Table.AddColumn(ColumnMergeOLD, ColumnMerge, each
                    if [Options.TMSName] = ALLMarker then
                        [Options.TMSName]
                    else
                        Record.Field(_,ColumnMerge&"OLD"), type text),

                    KeepValidTMS = Table.SelectRows(ColumnMergeNEW, each (Record.Field(_,ColumnMerge) <> null)),
                    RemoveColumns = Table.RemoveColumns(KeepValidTMS,{ColumnMerge&"OLD","WildcardMatch", "NewSourceColumn"})
                in
                    RemoveColumns,

            // Skapa en tabell med * raden lika många rader som CompleteTbl -1 (minus ALLMarker raden) och RowNumberALL för att kunna merga med CompleteTbl
            // ---------------------------------------------------------------------------------
            ALLTbl = 
                let
                    // Antal rader är lika med det antal rader som CompleteTbl innehåller minus 1 för ALLMarker raden
                    CompleteNumberOfRows = Table.RowCount(CompleteTbl) -1,
                    ALLRowOnly = Table.SelectRows(CompleteTbl, each ([Options.TMSName] = ALLMarker)),
                    AddRows = Table.AddColumn(ALLRowOnly, "RowNumberALL", each {1..CompleteNumberOfRows}),
                    ExpandRows = Table.ExpandListColumn(AddRows, "RowNumberALL"),
                    ChangedType = Table.TransformColumnTypes(ExpandRows,{{"RowNumberALL", Int64.Type}})
                in
                    ChangedType,

            // Börja med att ta bort ALLMarker raden och lägg till radnummer för att kunna merga med ALLTbl
            // ---------------------------------------------------------------------------------
            CompleteTblAdjusted = Table.SelectRows(CompleteTbl, each (Record.Field(_,ColumnMerge) <> ALLMarker)),
            RowNumberComplete = Table.AddIndexColumn(CompleteTblAdjusted, "RowNumberComplete", 1, 1),
            CompleteTblFinal = RowNumberComplete,

            // Merge nu CompleteTbl med ALLTbl för att få ALL värden bredvid
            // ---------------------------------------------------------------------------------
            MergeCompleteALL = Table.NestedJoin(CompleteTblFinal,{"RowNumberComplete"},ALLTbl,{"RowNumberALL"},"ALL",JoinKind.FullOuter),
            ExpandedCompleteALL = Table.ExpandTableColumn(MergeCompleteALL, "ALL", {"Options.TMSName", "Options.TOChoice", "Options.CommonChoice", "Options.ExclTechnicalCost", "Options.TMSVolume", "Options.OperatingCost", "Options.OperatingCostDepTime", "Options.CommonKeyValue", "Options.TOKeyValue", "Options.CommonCostDepTime", "Options.TOCostDepTime", "Options.OneTimeCostDepTime", "Options.PurchaseCostDepTime", "Options.InstallationCostDepTime"}, {"ALL.Options.TMSName", "ALL.Options.TOChoice", "ALL.Options.CommonChoice", "ALL.Options.ExclTechnicalCost", "ALL.Options.TMSVolume", "ALL.Options.OperatingCost", "ALL.Options.OperatingCostDepTime", "ALL.Options.CommonKeyValue", "ALL.Options.TOKeyValue", "ALL.Options.CommonCostDepTime", "ALL.Options.TOCostDepTime", "ALL.Options.OneTimeCostDepTime", "ALL.Options.PurchaseCostDepTime", "ALL.Options.InstallationCostDepTime"}),

            // Gå igenom samtliga kolumner. 
            // Saknas värde för den specifika raden (Options.X) och det finns värde i ALL.Options.X så används den senare
            // Annars används Options.X
            // ---------------------------------------------------------------------------------
            Option.TOChoice = Table.AddColumn(ExpandedCompleteALL, "Option.TOChoice", each
            if [Options.TOChoice] = null and [ALL.Options.TOChoice] <> null then
                [ALL.Options.TOChoice]
            else
                [Options.TOChoice], type text),

            Option.CommonChoice = Table.AddColumn(Option.TOChoice, "Option.CommonChoice", each
            if [Options.CommonChoice] = null and [ALL.Options.CommonChoice] <> null then
                [ALL.Options.CommonChoice]
            else
                [Options.CommonChoice], type text),

            Option.ExclTechnicalCost = Table.AddColumn(Option.CommonChoice, "Option.ExclTechnicalCost", each
            if [Options.ExclTechnicalCost] = null and [ALL.Options.ExclTechnicalCost] <> null then
                [ALL.Options.ExclTechnicalCost]
            else
                [Options.ExclTechnicalCost], type logical),

            Option.TMSVolume = Table.AddColumn(Option.ExclTechnicalCost , "Option.TMSVolume", each
            if [Options.TMSVolume] = null and [ALL.Options.TMSVolume] <> null then
                [ALL.Options.TMSVolume]
            else
                [Options.TMSVolume], type number),

            Option.OperatingCost = Table.AddColumn(Option.TMSVolume, "Option.OperatingCost", each
            if [Options.OperatingCost] = null and [ALL.Options.OperatingCost] <> null then
                [ALL.Options.OperatingCost]
            else
                [Options.OperatingCost], type number),

            Option.OperatingCostDepTime = Table.AddColumn(Option.OperatingCost, "Option.OperatingCostDepTime", each
            if [Options.OperatingCostDepTime] = null and [ALL.Options.OperatingCostDepTime] <> null then
                [ALL.Options.TMSVolume]
            else
                [Options.OperatingCostDepTime], type number),

            Option.CommonKeyValue = Table.AddColumn(Option.OperatingCostDepTime, "Option.CommonKeyValue", each
            if [Options.CommonKeyValue] = null and [ALL.Options.CommonKeyValue] <> null then
                [ALL.Options.CommonKeyValue]
            else
                [Options.CommonKeyValue], type number),

            Option.TOKeyValue =Table.AddColumn(Option.CommonKeyValue, "Option.TOKeyValue", each
            if [Options.TOKeyValue] = null and [ALL.Options.TOKeyValue] <> null then
                [ALL.Options.TOKeyValue]
            else
                [Options.TOKeyValue], type number),

            Option.CommonCostDepTime = Table.AddColumn(Option.TOKeyValue, "Option.CommonCostDepTime", each
            if [Options.CommonCostDepTime] = null and [ALL.Options.CommonCostDepTime] <> null then
                [ALL.Options.CommonCostDepTime]
            else
                [Options.CommonCostDepTime], type number),

            Option.TOCostDepTime = Table.AddColumn(Option.CommonCostDepTime, "Option.TOCostDepTime", each
            if [Options.TOCostDepTime] = null and [ALL.Options.TOCostDepTime] <> null then
                [ALL.Options.TOCostDepTime]
            else
                [Options.TOCostDepTime], type number),

            Option.OneTimeCostDepTime = Table.AddColumn(Option.TOCostDepTime, "Option.OneTimeCostDepTime", each
            if [Options.OneTimeCostDepTime] = null and [ALL.Options.OneTimeCostDepTime] <> null then
                [ALL.Options.OneTimeCostDepTime]
            else
                [Options.OneTimeCostDepTime], type number),

            Option.PurchaseCostDepTime = Table.AddColumn(Option.OneTimeCostDepTime, "Option.PurchaseCostDepTime", each
            if [Options.PurchaseCostDepTime] = null and [ALL.Options.PurchaseCostDepTime] <> null then
                [ALL.Options.PurchaseCostDepTime]
            else
                [Options.PurchaseCostDepTime], type number),

            Option.InstallationCostDepTime = Table.AddColumn(Option.PurchaseCostDepTime, "Option.InstallationCostDepTime", each
            if [Options.InstallationCostDepTime] = null and [ALL.Options.InstallationCostDepTime] <> null then
                [ALL.Options.InstallationCostDepTime]
            else
                [Options.InstallationCostDepTime], type number),

            // Ta bort arbetskolumner
            // ---------------------------------------------------------------------------------
            RemovedColumns2 = Table.RemoveColumns(Option.InstallationCostDepTime,{"ALL.Options.TMSName", "ALL.Options.TOChoice", "ALL.Options.CommonChoice", "ALL.Options.ExclTechnicalCost", "ALL.Options.TMSVolume", "ALL.Options.OperatingCost", "ALL.Options.OperatingCostDepTime", "ALL.Options.CommonKeyValue", "ALL.Options.TOKeyValue", "ALL.Options.CommonCostDepTime", "ALL.Options.TOCostDepTime", "ALL.Options.OneTimeCostDepTime", "ALL.Options.PurchaseCostDepTime", "ALL.Options.InstallationCostDepTime", "Options.TMSName", "Options.TOChoice", "Options.CommonChoice", "Options.ExclTechnicalCost", "Options.TMSVolume", "Options.CommonKeyValue", "Options.TOKeyValue", "Options.CommonCostDepTime", "Options.TOCostDepTime", "Options.OneTimeCostDepTime", "Options.PurchaseCostDepTime", "Options.InstallationCostDepTime", "Options.OperatingCost", "Options.OperatingCostDepTime", "RowNumberComplete"})
        in
            RemovedColumns2,

    // -----------------------------------------------------------------
    COST =
    // -----------------------------------------------------------------
        let
            Source = 
                let
                    Source = TableMerge,
                    DuplicatedColumn = Table.DuplicateColumn(Source, ColumnMerge, "MergeColumnTemp"),
                    Uppercase = Table.TransformColumns(DuplicatedColumn,{{"MergeColumnTemp", Text.Upper, type text}})
                in
                    Uppercase,

            Option = COSTTbl,

            MergeOption = Table.NestedJoin(Source,{"MergeColumnTemp"},Option,{"COSTBeskrivning"},"Option",JoinKind.FullOuter),
            ExpandOption = Table.ExpandTableColumn(MergeOption, "Option", {"ExclBudget", "ExclTechnicalCost", "ExclCommon", "ExclTO"}, {"Option.ExclBudget", "Option.ExclTechnicalCost", "Option.ExclCommon", "Option.ExclTO"}),
            ExcludeNULLCost = Table.SelectRows(ExpandOption, each ([MergeColumnTemp] <> null)),
            DEFAULTExcludeBudget = Table.ReplaceValue(ExcludeNULLCost,null,DEFAULT_ExcludeBudget,Replacer.ReplaceValue,{"Option.ExclBudget"}),
            DEFAULTExcludeTechnicalCost = Table.ReplaceValue(DEFAULTExcludeBudget,null,DEFAULT_ExcludeTechnicalCost,Replacer.ReplaceValue,{"Option.ExclTechnicalCost"}),
            DEFAULTExcludeCommon = Table.ReplaceValue(DEFAULTExcludeTechnicalCost,null,DEFAULT_ExcludeCommon,Replacer.ReplaceValue,{"Option.ExclCommon"}),
            DEFAULTExcludeTO = Table.ReplaceValue(DEFAULTExcludeCommon,null,DEFAULT_ExcludeTO,Replacer.ReplaceValue,{"Option.ExclTO"}),
            RemoveTempColumn = Table.RemoveColumns(DEFAULTExcludeTO,{"MergeColumnTemp"})
        in
            RemoveTempColumn,

    // -----------------------------------------------------------------
    TO_ALL =
    // -----------------------------------------------------------------
        let
            Source = TMSTbl,
            SelectALL = Table.SelectRows(Source, each ([TMSName] = ALLMarker)),
            Return = Text.Upper(SelectALL[TOChoice]{0})
        in
            Return,

    // -----------------------------------------------------------------
    COMMON_ALL =
    // -----------------------------------------------------------------
        let
            Source = TMSTbl,
            SelectALL = Table.SelectRows(Source, each ([TMSName] = ALLMarker)),
            Return = Text.Upper(SelectALL[CommonChoice]{0})
        in
            Return,

// =======================================================================
XXX_RESULT_XXX = null,
// =======================================================================

        Result = 
        if Choice = ChoiceTMSGROUPING then
            TMSGrouping
        else if Choice = ChoiceTMS then
            TMS
        else if Choice = ChoiceCOST then
            COST
        else if Choice = ChoiceTOALL then
            TO_ALL
        else if Choice = ChoiceCOMMONALL then
            COMMON_ALL
        else
            null

    in
        Result
///*
in
    fnOptions
//*/