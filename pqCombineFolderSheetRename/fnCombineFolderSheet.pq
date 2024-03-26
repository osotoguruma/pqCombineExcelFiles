let 
    CombineFolderSheet = (pFolder as text) =>
    let
        Source      = Excel.Workbook(File.Contents(pFolder), null, true),
        KindSheet   = Table.SelectRows(Source, each ([Kind] = "Sheet")),
        ToPromote   = Table.TransformColumns(KindSheet,{{"Data", each Table.PromoteHeaders(_), type table}}),
        ListTable   = ToPromote[Data],
        ListColumns = Table.ColumnNames(ListTable{0}),
        ToTableExpand     = Table.ExpandTableColumn(ToPromote, "Data", ListColumns, ListColumns),
        ToTable = Table.SelectColumns(ToTableExpand,ListColumns)
    in
        ToTable
in 
    CombineFolderSheet