let 
    fnCombineMultiTablesEXCEL = (
            pFolder                as text,
            optional pFolderFilter as any,
            optional pFilesFilter  as any,
            optional pFilterItem   as any,
            optional pColTechnic   as any,
            optional pSortField    as any
    ) as table =>

        let

            //--------------------------------------------------------------------------------------------------
            // Fonction permettant de chager le binaire d'un classeur Excel
            //--------------------------------------------------------------------------------------------------
            LoadContents = (ParamContents) => 
                let 
                    Src=Excel.Workbook(ParamContents, null, true)
                in 
                    Src,

            //--------------------------------------------------------------------------------------------------
            // Fonction permettant de chager le dossier 
            //--------------------------------------------------------------------------------------------------
            GetSource = () as any =>
                let 
                    GetSrc=try Folder.Files(pFolder) otherwise null
                in 
                    GetSrc,

            //--------------------------------------------------------------------------------------------------
            // Fonction permettant de construire la table
            //--------------------------------------------------------------------------------------------------
            Build = () as any =>
                let         

                    //--------------------------------------------------------------------------------------------------
                    // Filtre sur les noms de fichiers
                    //--------------------------------------------------------------------------------------------------
                    FilterSource = if pFilesFilter is null or pFilesFilter = false then 
                                    Source 
                                    else 
                                        //---------------------------------------------------------------------------------
                                        // Bloc de traitements quand plusieurs instructions sont nécessaires
                                        //---------------------------------------------------------------------------------
                                        let 
                                            //-----------------------------------------------------------------------------
                                            // On va en plus vérifie si la table est vide
                                            //-----------------------------------------------------------------------------
                                            // On remarque ici que dans une instruction if then else
                                            // on peut placer un bloc d'instructions via let et in
                                            //-----------------------------------------------------------------------------
                                            TblSource1=Table.SelectRows(Source, each Text.StartsWith([Name], pFilesFilter)),
                                            TblSource2=if Table.IsEmpty(TblSource1) then #table({},{}) else TblSource1
                                        in 
                                            TblSource2,

                    //--------------------------------------------------------------------------------------------------
                    // Suppression des colonnes techniques inutiles
                    //--------------------------------------------------------------------------------------------------
                    DelOtherColumns = if Table.IsEmpty(FilterSource) then 
                                         #table({},{})
                                         else Table.SelectColumns(FilterSource,{"Content", "Name", "Folder Path"}),

                    //--------------------------------------------------------------------------------------------------
                    // Cas des fichiers cachés
                    //--------------------------------------------------------------------------------------------------
                    FilterHidden = Table.SelectRows(DelOtherColumns, each [Attributes]?[Hidden]? <> true),    

                    //--------------------------------------------------------------------------------------------------
                    // Transformation des binaires
                    //--------------------------------------------------------------------------------------------------
                    CallFunction = Table.AddColumn(FilterHidden, "TransformFile", each LoadContents([Content]?)),

                    //--------------------------------------------------------------------------------------------------
                    // Extraction des binaires
                    //--------------------------------------------------------------------------------------------------
                    Expand1 = Table.ExpandTableColumn(CallFunction, "TransformFile", 
                                    {"Name", "Data", "Item", "Kind", "Hidden"}, 
                                    {"TableExcel", "Expand1.Data", "Expand1.Item", "Expand1.Kind", "Expand1.Hidden"}),

                    //--------------------------------------------------------------------------------------------------
                    // Filtre sur le dossier à retenir
                    //--------------------------------------------------------------------------------------------------
                    FilterToFolder = if pFolderFilter is null or pFolderFilter = false then 
                                        Expand1
                                        else Table.SelectRows(Expand1, each Text.StartsWith([Folder Path], pFolderFilter)),

                    //--------------------------------------------------------------------------------------------------
                    // Renommage des colonnes techniques qui seront conservées si besoin
                    //--------------------------------------------------------------------------------------------------
                    RenameColumnsExcel = Table.RenameColumns(FilterToFolder,{{"Name", "ClasseurExcel"}, {"Folder Path", "DossierExcel"}}),

                    //--------------------------------------------------------------------------------------------------
                    // Filtre sur le nom des tableaux structurés
                    //--------------------------------------------------------------------------------------------------
                    FilterLinesTable = if pFilterItem is null or pFilterItem = false then 
                                        Table.SelectRows(RenameColumnsExcel, each ([Expand1.Kind] = "Table"))
                                        else let 
                                                    //------------------------------------------------------------------------------
                                                    // Procédure de Filtrage sur le Kind et Item - bloc d'instrcutions
                                                    //------------------------------------------------------------------------------
                                                    SelKind=Table.SelectRows(RenameColumnsExcel, each ([Expand1.Kind] = "Table")),
                                                    SelItem=Table.SelectRows(RenameColumnsExcel, each (Text.StartsWith([Expand1.Item], pFilterItem)))
                                                in 
                                                    SelItem,                                                    

                    //--------------------------------------------------------------------------------------------------
                    // Extraction des colonnes titre à partir de la 1ère table
                    //--------------------------------------------------------------------------------------------------
                    FirstTable=FilterLinesTable[Expand1.Data]{0},
                    ListColumns=Table.ColumnNames(FirstTable),

                    //--------------------------------------------------------------------------------------------------
                    // On ne retient que les colonnes concernées
                    //--------------------------------------------------------------------------------------------------
                    Expand2 = Table.ExpandTableColumn(FilterLinesTable, "Expand1.Data", ListColumns),

                    //--------------------------------------------------------------------------------------------------
                    // La colonne des binaires n'est plus utile
                    //--------------------------------------------------------------------------------------------------
                    DeleteBinaryContents = Table.RemoveColumns(Expand2,{"Content"}),

                    //--------------------------------------------------------------------------------------------------
                    // Transformation brute avec ajout ou non des colonnes techniques
                    //--------------------------------------------------------------------------------------------------
                    ToTable = if pColTechnic is null or pColTechnic = false then 
                                Table.SelectColumns(DeleteBinaryContents,ListColumns)
                                else Table.SelectColumns(DeleteBinaryContents,{"ClasseurExcel", "DossierExcel", "TableExcel"} & ListColumns),

                    //--------------------------------------------------------------------------------------------------
                    // Traitement si nécessaire des critères de tri
                    //--------------------------------------------------------------------------------------------------
                    SortTable = if pSortField is null or pSortField = false then 
                                ToTable
                                else let
                                            //--------------------------------------------------------------------------
                                            // Procédure de tri - bloc d'instructions
                                            //--------------------------------------------------------------------------
                                            pOrder=List.Repeat({0},List.Count(Text.Split(pSortField,";"))),
                                            pCrit=List.Zip({Text.Split(pSortField,";"),pOrder}),    
                                            ToSort=Table.Sort(ToTable,pCrit)
                                        in 
                                            ToSort
                in 
                    SortTable,

            //--------------------------------------------------------------------------------------------------
            // MAIN PROCEDURE ET APPEL DES FONCTIONS
            //--------------------------------------------------------------------------------------------------
            Source = GetSource(),
            ToBuild=if Source is null then 
                    #table ({},{})
                    else Build()

        in
            ToBuild

in 
    fnCombineMultiTablesEXCEL