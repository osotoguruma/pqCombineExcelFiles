let 
    CombineFolderSheetRenameColumns = (pFolder as text) as table =>
    let

        //------------------------------------------------------------------------------------------------
        // Combinaison des onglets d'un fichier Excel externe
        // avec renommage et typage des colonnes
        //------------------------------------------------------------------------------------------------
        
        //------------------------------------------------------------------------------------------------
        // On récupère les paramètres
        //------------------------------------------------------------------------------------------------

        // Le classeur avec ses onglets
        FilePath = fnGetParameters("TB_PARAMS","DOSSIER") & "\" & fnGetParameters("TB_PARAMS","FICHIER"),

        // La table de mapping avec les nouveaux types
        T_Mapping=fnGetParameters("TB_PARAMS","TABLE_RENAME"),

        //------------------------------------------------------------------------------------------------
        // Table de mapping
        //------------------------------------------------------------------------------------------------
        TbMapping = Excel.CurrentWorkbook(){[Name=T_Mapping]}[Content],

        //------------------------------------------------------------------------------------------------
        // On prépare la table de mapping pour typer les colonnes
        //------------------------------------------------------------------------------------------------
        ListColumnsTbMapping=Table.ColumnNames(TbMapping),
        Ancien=ListColumnsTbMapping{0},
        Nouveau=ListColumnsTbMapping{1},
        Type=ListColumnsTbMapping{2},

        //------------------------------------------------------------------------------------------------
        // On prépare le mapping des noms de colonnes
        //------------------------------------------------------------------------------------------------
        //ToStringAncien="TbMapping[" & Ancien & "]",
        //ListAncien=Expression.Evaluate(ToStringAncien,[TbMapping=TbMapping]),

        ListAncien=Table.Column(TbMapping,Ancien),

        //ToStringNouveau="TbMapping[" & Nouveau & "]",
        //ListNouveau=Expression.Evaluate(ToStringNouveau,[TbMapping=TbMapping]),
        
        ListNouveau=Table.Column(TbMapping,Nouveau),
        
        //ToStringType="TbMapping[" & Type & "]",
        //ListType=Expression.Evaluate(ToStringType,[TbMapping=TbMapping]),
        
        ListType=Table.Column(TbMapping,Type),

        //------------------------------------------------------------------------------------------------
        // On applique le renommage des colonnes
        ChangeColTypes = Table.TransformColumnTypes(TbMapping,{{ListColumnsTbMapping{0}, type text}, {ListColumnsTbMapping{1}, type text}, {ListColumnsTbMapping{2}, type text}}),

        //------------------------------------------------------------------------------------------------
        // Suppression de la colonne ancien dans la table de mapping pour attérir sur le nouveau nom
        //------------------------------------------------------------------------------------------------
        ListAfterRename=Table.ColumnNames(ChangeColTypes),
        SupprColAncien = Table.RemoveColumns(ChangeColTypes,{ListAfterRename{0}}),

        //------------------------------------------------------------------------------------------------
        // Construction du contexte powerquery pour le typage des colonnes
        //------------------------------------------------------------------------------------------------
        ToType = Table.TransformColumns(SupprColAncien,{{ListColumnsTbMapping{2}, each Expression.Evaluate(_, [Currency.Type=Currency.Type, Int64.Type=Int64.Type, Percentage.Type=Percentage.Type])}}),

        //------------------------------------------------------------------------------------------------
        // Traitements sur les colonnes pour le typage
        //------------------------------------------------------------------------------------------------
        ToField = Table.AddColumn(ToType, "CustomTypes", each Record.FieldValues(_)),
        RemoveOtherCols = Table.SelectColumns(ToField,{"CustomTypes"}),
        ToListType = RemoveOtherCols[CustomTypes],
        
        //------------------------------------------------------------------------------------------------
        // Capture des onglets
        //------------------------------------------------------------------------------------------------
        Source = Excel.Workbook(File.Contents(FilePath), null, true),
        KindSheet = Table.SelectRows(Source, each ([Kind] = "Sheet")),
        TransformColumns_TABLE = Table.TransformColumns(KindSheet,{{"Data", each Table.PromoteHeaders(_), type table}}),

        //------------------------------------------------------------------------------------------------
        // Renommage des colonnes
        //------------------------------------------------------------------------------------------------
        TransformColumns_ZIP = Table.TransformColumns(TransformColumns_TABLE, {{"Data",  each Table.RenameColumns(_, List.Zip({ListAncien,ListNouveau}) , MissingField.Ignore), type table}}),
        
        //------------------------------------------------------------------------------------------------
        // Transformation en table de données
        //------------------------------------------------------------------------------------------------
        Columns_DATA = Table.SelectColumns(TransformColumns_ZIP,{"Data"}),
        ToTable = Table.ExpandTableColumn(Columns_DATA, "Data", ListNouveau, ListNouveau),

        //------------------------------------------------------------------------------------------------
        // Application des types sur les colonnes
        //------------------------------------------------------------------------------------------------
        ToTableType = Table.TransformColumnTypes(ToTable,ToListType)  
        
    in
        ToTableType
in 
    CombineFolderSheetRenameColumns