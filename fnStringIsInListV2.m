// =================================================================================================================================
// fnStringIsInListV2
// Returns true if String exists in any of the KewordList items (partly or entire string)
// ---------------------------------------------------------------------------------------------------------------------------------
//  v1      2018-10-15  tswm00  New PQ
//  v2      2020-04-16  tswm00  Change to multiple options (textbased). Not backward compatible
// =================================================================================================================================

let
// =================================================================================================================================
// Return options
// =================================================================================================================================
OptionReturnTrueFalse = "LOGICAL",    // Returns true if match
OptionReturnPosition = "POSITION",    // Returns list position if match
OptionReturnValue = "VALUE",          // Returns matching value from keywordlist
// =================================================================================================================================
// Header
// =================================================================================================================================
///*
    fnStringInList =
        (
        SearchString_p            as text,                // Text to search for in list
        KeywordList_p             as list,                // List of textstrings to search within (partly or entire string)
        ReturnOption_p            as text                 // See above
        )
        as any =>

    let
//*/
// =================================================================================================================================
// TESTDATA
// =================================================================================================================================
/*
    SearchTable1 = Table.FromRecords
    ({
    [TMSName = "CONNECTION100",    Data1 = 0,    Data2 = "A"],
    [TMSName = "SWITCH100",        Data1 = 1,    Data2 = "B"],
    [TMSName = "TMS1",             Data1 = 1,    Data2 = "B"],
    [TMSName = "TMS2",             Data1 = 2,    Data2 = "C"]
    }),

    KeywordTable1 = Table.FromRecords
    ({
    [Keyword = "CONNECTION",    Data1 = 0,    Data2 = "A"],
    [Keyword = "SWITCH",        Data1 = 0,    Data2 = "A"]
    }),

    SearchString_p = SearchTable1[TMSName]{0},
    KeywordList_p = KeywordTable1[Keyword],
    ReturnOption_p = "value",
*/
// =================================================================================================================================
// Constants
// =================================================================================================================================
    SearchString = SearchString_p,
    KeywordList = List.Buffer(KeywordList_p), 
    ReturnOption = Text.Upper(ReturnOption_p),

// =================================================================================================================================
// Main function
// =================================================================================================================================

        //check if values in Keywordlist is in SearchString
        MatchFound = List.Transform(KeywordList, each Text.Contains(SearchString, _, Comparer.OrdinalIgnoreCase)), 

        PositionToReturn = List.PositionOf(MatchFound, true),

        ValueToReturn = 
        if PositionToReturn >= 0 then
            KeywordList{PositionToReturn}
        else
            null,

         LogicalToReturn = List.PositionOf(MatchFound, true) >= 0,

        //Return either position of true/false
        Return =
        if ReturnOption = OptionReturnTrueFalse then
            LogicalToReturn
        else if ReturnOption = OptionReturnPosition then
            PositionToReturn
        else if ReturnOption = OptionReturnValue then
            ValueToReturn
        else
            null
    in
        Return
///*
in
    fnStringInList
//*/