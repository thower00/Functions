let
// -------------------------------------------------------
pqName = "fnCountIFV1",
// -------------------------------------------------------
// -------------------------------------------------------
/* Beskrivning
*/
// -------------------------------------------------------
// Ändringslogg
// -------------------------------------------------------
/*
2021-02-16    tswm00    V1: New PQ
*/
// -------------------------------------------------------

  countif = 
    (
                tbl             as table,
                col             as text,
                value             as any
    )
    as number =>

    let
      select_rows = Table.SelectRows(tbl, each Record.Field(_, col) = value),
      count_rows = Table.RowCount(select_rows)
    in
      count_rows
in
    countif