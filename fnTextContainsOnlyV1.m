let
// =================================================================================================================================
// fnTextContainsOnlyV1
// Checks if string contains only values fromlist
// ---------------------------------------------------------------------------------------------------------------------------------
//  V1      2018-10-25  tswm00  New function
// =================================================================================================================================

    fnTextContainsOnly =
    (
                    CheckString         as text,            // String to check
                    ValueList           as list,            // List of values to check against
        optional    CaseSensitive       as nullable logical // If omitted default is false (not case sensitive)
    )
    as logical => 
    
    let
// =================================================================================================================================
// Constants
// =================================================================================================================================

// *********************************************************************************************************************************
// Main query
// *********************************************************************************************************************************

    IsCaseSensitive =
    if CaseSensitive = null then    
        false
    else
        CaseSensitive,
        
    Result =
    if CaseSensitive then
        Text.Length(CheckString) = Text.Length(Text.Select(CheckString,ValueList))
    else
        Text.Length(Text.Upper(CheckString)) = Text.Length(Text.Select(Text.Upper(CheckString),List.Transform(ValueList, each Text.Upper(_))))
    in
        Result

in
    fnTextContainsOnly