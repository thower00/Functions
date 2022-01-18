let
// =================================================================================================================================
// fnValidateSiteIDV1
// Checks if SiteID matches specified pattern (e.g. CCCCNNNMNN)
// ---------------------------------------------------------------------------------------------------------------------------------
//  V1      2019-10-08  tswm00  New function
// =================================================================================================================================
/*
// DEBUG
    SiteID = Text.Upper("SITE123"),
    Pattern = Text.Upper("CCCCNNNMNN"),
*/
    fnValidateSiteID =
    (
                    SiteID    as text,            // SiteID to check
                    Pattern   as text            // Pattern to check against (C for Character, N for number, M for Misc)
    )
    as logical => 
    
    let
// =================================================================================================================================
// Constants
// =================================================================================================================================
        charList =  List.Buffer(List.Transform({65..90}, each Character.FromNumber(_))),
        numList = List.Buffer(List.Transform({48..57}, each Character.FromNumber(_))),
        miscList = List.Buffer(List.Transform({35}, each Character.FromNumber(_))),

// *********************************************************************************************************************************
// Main query
// *********************************************************************************************************************************
        SiteIDLength = Text.Length(SiteID),
        PatternLength = Text.Length(Pattern),

        Check =
        if SiteIDLength = PatternLength then
            let
                Test0 =
                if SiteIDLength = PatternLength and SiteIDLength >= 1 then
                    if List.Contains(charList,Text.Range(SiteID,0,1)) and Text.Range(Pattern,0,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,0,1)) and Text.Range(Pattern,0,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,0,1)) and Text.Range(Pattern,0,1) = "M" then
                        true
                    else
                        false
                else
                    true,
                    
                Test1 =
                if SiteIDLength = PatternLength and SiteIDLength >= 2 then
                    if List.Contains(charList,Text.Range(SiteID,1,1)) and Text.Range(Pattern,1,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,1,1)) and Text.Range(Pattern,1,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,1,1)) and Text.Range(Pattern,1,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test2 =
                if SiteIDLength = PatternLength and SiteIDLength >= 3 then
                    if List.Contains(charList,Text.Range(SiteID,2,1)) and Text.Range(Pattern,2,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,2,1)) and Text.Range(Pattern,2,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,2,1)) and Text.Range(Pattern,2,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test3 =
                if SiteIDLength = PatternLength and SiteIDLength >= 4 then
                    if List.Contains(charList,Text.Range(SiteID,3,1)) and Text.Range(Pattern,3,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,3,1)) and Text.Range(Pattern,3,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,3,1)) and Text.Range(Pattern,3,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test4 =
                if SiteIDLength = PatternLength and SiteIDLength >= 5 then
                    if List.Contains(charList,Text.Range(SiteID,4,1)) and Text.Range(Pattern,4,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,4,1)) and Text.Range(Pattern,4,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,4,1)) and Text.Range(Pattern,4,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test5 =
                if SiteIDLength = PatternLength and SiteIDLength >= 6 then
                    if List.Contains(charList,Text.Range(SiteID,5,1)) and Text.Range(Pattern,5,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,5,1)) and Text.Range(Pattern,5,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,5,1)) and Text.Range(Pattern,5,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test6 =
                if SiteIDLength = PatternLength and SiteIDLength >= 7 then
                    if List.Contains(charList,Text.Range(SiteID,6,1)) and Text.Range(Pattern,6,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,6,1)) and Text.Range(Pattern,6,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,6,1)) and Text.Range(Pattern,6,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test7 =
                if SiteIDLength = PatternLength and SiteIDLength >= 8 then
                    if List.Contains(charList,Text.Range(SiteID,7,1)) and Text.Range(Pattern,7,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,7,1)) and Text.Range(Pattern,7,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,7,1)) and Text.Range(Pattern,7,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test8 =
                if SiteIDLength = PatternLength and SiteIDLength >= 9 then
                    if List.Contains(charList,Text.Range(SiteID,8,1)) and Text.Range(Pattern,8,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,8,1)) and Text.Range(Pattern,8,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,8,1)) and Text.Range(Pattern,8,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test9 =
                if SiteIDLength = PatternLength and SiteIDLength >= 10 then
                    if List.Contains(charList,Text.Range(SiteID,9,1)) and Text.Range(Pattern,9,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,9,1)) and Text.Range(Pattern,9,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,9,1)) and Text.Range(Pattern,9,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                Test10 =
                if SiteIDLength = PatternLength and SiteIDLength >= 11 then
                    if List.Contains(charList,Text.Range(SiteID,10,1)) and Text.Range(Pattern,10,1) = "C" then
                        true
                    else if List.Contains(numList,Text.Range(SiteID,10,1)) and Text.Range(Pattern,10,1) = "N" then
                        true
                    else if List.Contains(miscList,Text.Range(SiteID,10,1)) and Text.Range(Pattern,10,1) = "M" then
                        true
                    else
                        false
                else
                    true,

                PatternMatch = Test0 and Test1 and Test2 and Test3 and Test4 and Test5 and Test6 and Test7 and Test8 and Test9 and Test10
            in
                PatternMatch
        else
            false

    in
        Check

in
    fnValidateSiteID