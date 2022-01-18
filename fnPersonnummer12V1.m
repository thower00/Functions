let
// -------------------------------------------------------
// fnPersonnummer12V1
// Returnerar ett tolvsiffrigt personnummer utan bindestreck
// baserat på ett 10 eller 12-siffrigt personer med eller utan bindestreck
// -------------------------------------------------------
// 2018-01-31    tswm00    Ny PQ


// -------------------------------------------------------
    fnPersonnummer12 = (PNR as text) as text => 
// -------------------------------------------------------
let
    SeparatorPos = Text.PositionOf(PNR,"-"),
    YY = Number.From(Text.Middle(PNR, 0, 2)),

    PNRmod =
    if SeparatorPos <> -1 then
        // ta bort bindestreck
        Text.Middle(PNR,0,SeparatorPos) & Text.Middle(PNR,SeparatorPos+1,Text.Length(PNR))
    else
        PNR,

    TextLength = Text.Length(PNRmod),


    PersonID = 
    if TextLength = 12 then
        // Detta är ett 12-siffrigt personnummer utan bindestreck så bara returnera
        PNRmod
    else if TextLength = 10 then
        // Detta är ett 10 siffrigt pnr med bindestreck
        if TextLength = 7 then
            "20000" & PNRmod
        else
            if TextLength = 8 then
                "2000" & PNRmod
            else
                if TextLength = 9 then
                    "200" & PNRmod
                else
                    if TextLength = 10 and YY < 30 then
                        // Year (YY) är mindre än 30 så anta 2000-talet
                        "20" & PNRmod
                    else
                        // 1900-talet
                        "19" & PNRmod
    else
        "#N/A"

in
    PersonID

in
    fnPersonnummer12