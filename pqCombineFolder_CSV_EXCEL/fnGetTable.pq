let 
    //-------------------------------------------------------------------
    // Fonction permettant d'importer un tableau structuré
    // Ce tableau structuré doit exister au préalable
    //-------------------------------------------------------------------
    
    fnGetTable=(pTable as text) as table =>

    let
        Source = Excel.CurrentWorkbook(){[Name=pTable]}[Content]
    in
        try Source otherwise #table({"EMPTY"},{{pTable & " inconnue ou erreur"}})
in 
    fnGetTable