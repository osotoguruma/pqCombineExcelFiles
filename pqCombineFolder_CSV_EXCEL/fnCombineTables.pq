let 
    //-------------------------------------------------------------------------------
    // Combinaison de plusieurs tables à partir du classeur Excel courant
    // avec possibilité de filtrer sur le nom des tables (tableaux structurés)
    // au préalable les tableaux structurés doivent être créés
    // il est recommandé qu'ils commencent tous par le même préfixe
    //-------------------------------------------------------------------------------

    fnCombineAllTables = (pFilterTbl as text) =>
    
    let
        //---------------------------------------------------------------------------
        // On récupère les objets du classeur
        //---------------------------------------------------------------------------
        Source = Excel.CurrentWorkbook(),    

        //---------------------------------------------------------------------------
        // On applique le filtre sur les tableaux structurés
        //---------------------------------------------------------------------------
        FilterTable = if pFilterTbl is null then Source else Table.SelectRows(Source, each Text.StartsWith([Name], pFilterTbl)),

        //---------------------------------------------------------------------------
        // On récupère les tableaux structurés (tables)
        //---------------------------------------------------------------------------
        Content = FilterTable[Content],

        //---------------------------------------------------------------------------
        // Conversion en table d'objets
        //---------------------------------------------------------------------------
        ToTable = Table.FromList(Content, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

        //---------------------------------------------------------------------------
        // On se base sur le 1er tableau pour récupérer les titres de colonnes
        //---------------------------------------------------------------------------
        TblColumns = ToTable{0}[Column1],

        //---------------------------------------------------------------------------
        // Les colonnes sont transformées en liste
        //---------------------------------------------------------------------------
        LstColumns=Table.ColumnNames(TblColumns),

        //---------------------------------------------------------------------------
        // Création de la table contenant tous les tableaux structurés
        //---------------------------------------------------------------------------
        ExpandTable = Table.ExpandTableColumn(ToTable, "Column1", LstColumns, LstColumns)

    in
        ExpandTable
in 
    fnCombineAllTables