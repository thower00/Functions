// -------------------------------------------------------
pqName = "fnTransposeRepeatingRowsToColumnsV1",
// -------------------------------------------------------
// -------------------------------------------------------
// Beskrivning
// Läser in tabell med två kolumner där den vänstra kolumnen innehåller kolumnnamn och den högra kolumnvärden
// -------------------------------------------------------
/*
Parametrar:
IndataTbl_p             Tabell som skall transponeras
Indexer_p               Värde (oftast 0) där transponeringen skall börja
PatternCount_p          Hur många kolumner som skall transponeras till
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2020-10-02    tswm00    New PQ. 
*/
let
    Source = (
                IndataTbl_p     as table,
                Indexer_p       as number,
                PatternCount_p  as number
             ) => 
    let
        IndataTbl =     IndataTbl_p,
        Indexer =       Indexer_p,
        PatternCount =  PatternCount_p,
    
        NewTable = Table.Range(IndataTbl,Indexer,PatternCount),
        TransposedTable = Table.Transpose(NewTable),
        PromotedHeaders = Table.PromoteHeaders(TransposedTable),
        NextTable = if Indexer + PatternCount < Table.RowCount(IndataTbl)
            then
                Table.Combine({PromotedHeaders,@Source(IndataTbl, Indexer + PatternCount, PatternCount)})                
            else
                PromotedHeaders

    in
        NextTable
in
    Source