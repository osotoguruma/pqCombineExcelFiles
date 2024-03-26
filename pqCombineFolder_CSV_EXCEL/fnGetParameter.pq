(pTable as text, pName as any) =>
let
    ParamSource = Excel.CurrentWorkbook(){[Name=pTable]}[Content],
    ParamRow = Table.SelectRows(ParamSource, each ([PARAMETRE] = pName)),
    Value=
    if Table.IsEmpty(ParamRow)=true
        then null
    else Record.Field(ParamRow{0},"VALEUR")
in
    Value