// -------------------------------------------------------
// fnPowerTrim
// Trimmar bort 'char_to_trim' från 'text'
// Om 'trim_all' = true så trimmas alla förekomster av 'char_to_trim' bort
// Om 'trim_all' = false så trimmas 'char_to_trim' bort i början/slutet av strängen samt duplicerade förekomster i strängen. 

// -------------------------------------------------------
// 2018-08-13    tswm00    Ny PQ

// -------------------------------------------------------
// External references
// -------------------------------------------------------

// -------------------------------------------------------
// Query body
// -------------------------------------------------------

    (
    text as text,
    trim_all as logical,
    optional char_to_trim as text)
    =>

let
    char = if char_to_trim = null then " " else char_to_trim,
    split = Text.Split(text, char),
    removeblanks = List.Select(split, each _ <> ""),

    result=
    if trim_all then
        Text.Combine(removeblanks)
    else
        Text.Combine(removeblanks,char)
in
    result