// =================================================================================================================================
// fnStringIsInListV1
// Returns true if String exists in any of the KewordList items (partly or entire string)
// ---------------------------------------------------------------------------------------------------------------------------------
//  v1      2018-10-15  tswm00  New PQ
// =================================================================================================================================

let
    fnStringInList =
        (
        SearchString            as text,                // Text to search for in list
        KeywordList             as list,                // List of textstrings to search within (partly or entire string)
        optional ReturnPosition as nullable logical     // If true return position, else return exists (true/false). Default = false
        )
        as any =>

    let
// =================================================================================================================================
// Constants
// =================================================================================================================================
    ReturnPositionDefault = false,

// =================================================================================================================================
// Main function
// =================================================================================================================================
        ReturnPositionDefined =
        if ReturnPosition = null then
            ReturnPositionDefault
        else
            ReturnPosition,

        //check if values in Keywordlist is in SearchString
        MatchFound = List.Transform(List.Buffer(KeywordList), each Text.Contains(SearchString, _, Comparer.OrdinalIgnoreCase)), 

        //Return either position of true/false
        Return =
        if ReturnPositionDefined then
            List.PositionOf(MatchFound, true)
        else
            List.PositionOf(MatchFound, true) >= 0
    in
        Return
in
    fnStringInList