let
// -------------------------------------------------------
// fnMatchPatternV1
// -------------------------------------------------------
// Matches XXXXXXXXXXXXXXXXXx
// Pattern:
//  *XX*    Text contains
//  *XX     Text ends with
//  XX*     Text ends with
//  *       Any value
//  XX      Exact match

// -------------------------------------------------------
// 2018-12-03    tswm00    New PQ

// -------------------------------------------------------
// External references
// -------------------------------------------------------

// -------------------------------------------------------
// Constants
// -------------------------------------------------------
    FolderDelimiter = "\",
    FileExtensionDelimiter = ".",
    ErrorMsg = "#ERROR: File does not exist",

// -------------------------------------------------------
// Query body
// -------------------------------------------------------

    fnMatchPattern = 
        (
                    patternText_p   as text,
                    textToMatch_p   as text
        optional    ignoreCase_p    as nullable logical
        )
    as logical =>
	
    let
// -------------------------------------------------------
// Fix input values
// -------------------------------------------------------
        ignoreCase =
        if ignoreCase_p = null then
            false
        else
            ignoreCase_p
        
        patternText =
        if ignoreCase then
            Text.Upper(patternText_p)
        else
            patternText_p

        textToMatch =
        if ignoreCase then
            Text.Upper(textToMatch_p)
        else
            textToMatch_p
    
// -------------------------------------------------------
// Query body
// -------------------------------------------------------
        Result =
        if Text.StartsWith(textToMatch, "*") and Text.EndsWith(textToMatch, "*") then
            Text.Contains(patternText, Text.Remove(textToMatch, "*"))
        else if Text.StartsWith(textToMatch, "*") then
            Text.EndsWith(patternText, Text.Remove(textToMatch, "*"))
        else if Text.EndsWith(textToMatch, "*") then
            Text.StartsWith(patternText, Text.Remove(textToMatch, "*")) 
        else if textToMatch = "*" then
            true
        else if textToMatch = patternText then
            true
        else
            false    
    in
        Result
in
    fnFileFromFolder