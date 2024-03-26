let fnCombineExcelFiles=(pFolder as text, optional pFilter as text) as table =>
    //-----------------------------------------------------------------------------------
    // Importation de fichiers Excel contenus dans un dossier
    //-----------------------------------------------------------------------------------
    let
        FIELD_CUSTOM="ContentField",

        //-----------------------------------------------------------------------------------
        // Dossier des fichiers Excel
        //-----------------------------------------------------------------------------------
        Source = Folder.Files(pFolder),

        //-----------------------------------------------------------------------------------
        // Eventuellement on peut filtrer sur les noms de fichiers
        //-----------------------------------------------------------------------------------
        Filter = if pFilter is null then 
                    Source
                    else Table.SelectRows(Source, each Text.StartsWith([Name], pFilter)),

        //-----------------------------------------------------------------------------------
        // Pour le renommage des noms de colonnes techniques
        //-----------------------------------------------------------------------------------
        Lst_1={"Name", "Data", "Item", "Kind", "Hidden"},
        Lst_2={"Name.bis", "Data", "Item", "Kind", "Hidden"},

        //-----------------------------------------------------------------------------------
        // On récupère les données binaires dans une colonne spécifique
        //-----------------------------------------------------------------------------------
        AddCustomContentField = Table.AddColumn(Filter, FIELD_CUSTOM, each Excel.Workbook([Content], true)),

        //-----------------------------------------------------------------------------------
        // Renommage des colonnes techniques
        //-----------------------------------------------------------------------------------
        ExpandCustomContentField = Table.ExpandTableColumn(AddCustomContentField, FIELD_CUSTOM, Lst_1, Lst_2),    

        //-----------------------------------------------------------------------------------
        // On déploie les données binaires des fichiers Excel
        //-----------------------------------------------------------------------------------
        ExpandToTable=Table.Combine(ExpandCustomContentField[Data])
        
    in
        ExpandToTable
in 
    fnCombineExcelFiles