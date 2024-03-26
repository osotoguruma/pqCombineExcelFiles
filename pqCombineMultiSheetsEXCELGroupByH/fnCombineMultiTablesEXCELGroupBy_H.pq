let 

    //-------------------------------------------------------------------------------------------------------------------------------------
    // Fonction permettant de combiner plusieurs classeurs Excel avec plusieurs onglets
    // Recommandations :
    // Les fichiers Excel doivent avoir tous la même structure
    // Les données doivent être organisées dans des tableaux structurés
    // Pour des besoins de filtrages éventuels sur les noms de TS il est recommandé qu'ils débutent tous par le même préfixe
    // Exemple de structure
    // IDPROD     Janv    Fev     Mars    Avr     Mai     Juin    Juil    Aout    Sep     Oct     Novb      Decb
    // P0001      105     190     107     100     1580    250     300     256     147     123     856       145
    // Dans ce cas pRightOffset = 1
    // pSortColumn = "IDPROD"
    // pGroupByColumn = "IDPROD"
    // pFilesFilter : permet de filtrer sur les noms de fichiers dans le dossier y compris les dossiers
    // pFilter : permet de filtrer sur les tables (noms des tableaux structurés)
    //
    // Autre exemple
    // -------------
    // IDPROD   LIBELLE        Janv    Fev     Mars    Avr     Mai     Juin    Juil    Aout    Sep     Oct     Novb      Decb
    // P0001    LIB1           105     190     107     100     1580    250     300     256     147     123     856       145
    // Dans ce cas pRightOffset = 2 (pour prendre compte IDPROD et LIBELLE)
    // pSortColumn = "IDPROD"
    // pGroupByColumn = "IDPROD;LIBELLE" --> cela va constituer une liste de colonnes - ";" est le séparateur imposé
    // pRightOffset : correspond donc au nombre de colonnes titres à reprendre dans le GroupBy
    //-------------------------------------------------------------------------------------------------------------------------------------

    fnCombineMultiTablesEXCELGroupBy_H =(
        pFolder as text,
        pRightOffset as number,
        pKind as text,
        pSortColumn as text,
        pGroupByColumn as text,
        optional pFilesFilter as text,
        optional pFilter as text
    ) as table =>

    let
        //-----------------------------------------------------------------------------
        // On récupère le contenu du dossier
        //-----------------------------------------------------------------------------
        Source = Folder.Files(pFolder),
        FilterSource = if pFilesFilter is null then 
                        Source 
                        else 
                            let 
                                TblSource1=Table.SelectRows(Source, each Text.StartsWith([Name], pFilesFilter)),
                                TblSource2=if Table.IsEmpty(TblSource1) then null else TblSource1
                            in 
                                TblSource2,

        //---------------------------------------------------------------------------------------------------
        // Fonction permettant de construire la table
        // L'appel de cette fonction permet de sortir de la requête proprement si aucune donnée à traiter
        // Cette fonction joue un peu le rôle d'une fonction ou sub en VBA
        //---------------------------------------------------------------------------------------------------
        BuildTable = (Content as any) =>
            let 
                //-----------------------------------------------------------------------------
                // Pour chaque binaire on récupère les tables (tableaux structurés)
                // On fait appel à une boucle pour traiter la liste des binaires à fusionner
                //-----------------------------------------------------------------------------
                TbListCombine = List.Accumulate(
                                    Content,          // Liste des binaires à combiner
                                    #table({},{}),    // Initialisation avec une table vide
                                    (state,current) =>  // Lancement de la boucle
                                            let 
                                                Classeur=Excel.Workbook(current), // On récupère le binaire
                                                Filter=Table.SelectRows(Classeur, each ([Kind] = pKind)), // On retient le type désiré (Table)
                                                CombineFilter=Table.Combine({state,Filter}) // Les binaires sont mis bout à bout
                                            in 
                                                CombineFilter                               
                                ),

                //-----------------------------------------------------------------------------
                // Doit-on appliquer un filtre sur les tables à retenir                 
                //-----------------------------------------------------------------------------
                FilterTable = if pFilter is null then 
                                TbListCombine 
                                else Table.SelectRows(TbListCombine, each Text.StartsWith([Item], pFilter)),   

                //-----------------------------------------------------------------------------
                // On extrait le les colonnes du 1er classeur qui va servir de référence 
                //-----------------------------------------------------------------------------
                ListData=FilterTable[Data],
                ListColumns=Table.ColumnNames(ListData{0}),

                //-----------------------------------------------------------------------------
                // On récupère les data dans les binaires pour les convertir en table
                //-----------------------------------------------------------------------------
                ToData = Table.ExpandTableColumn(FilterTable, "Data", ListColumns),

                //-----------------------------------------------------------------------------
                // On ne retient que les colonnes de référence
                //-----------------------------------------------------------------------------
                ToSelectColumns = Table.SelectColumns(ToData,ListColumns),
                
                //-----------------------------------------------------------------------------
                // Tri selon les n premières colonnes de gauche retenues
                //-----------------------------------------------------------------------------
                Sort = Table.Sort(ToSelectColumns,{{pSortColumn, Order.Ascending}}),

                //-----------------------------------------------------------------------------
                // Construction des totaux par colonne selon le regroupement
                // Pour la totalisation des colonnes on passe à nouveau par une boucle
                //-----------------------------------------------------------------------------
                ListRight=List.LastN(ListColumns,List.Count(ListColumns)-pRightOffset),
                ToStringTotal=List.Accumulate(
                                        ListRight, // Liste des colonnes à cumuler
                                        "",        // Initialisation de l'accumulateur
                                        (state,current) => // Lancement de la boucle
                                            let 
                                                acc=if state = "" then "" else state & ", ", // pour éviter une virgule de trop en tête
                                                // L'accumulateur sous forme de chaine pour l'évaluer en expression PowerQuery à l'issue
                                                Tot=acc & "{""" & current & """,  each List.Sum([" & current & "]), type number}"                                                                  
                                            in 
                                                Tot
                            ),

                //-----------------------------------------------------------------------------
                // Evaluation pour application dans le regroupement des totaux
                //-----------------------------------------------------------------------------
                ToEvalTotal=Expression.Evaluate("{" & ToStringTotal & "}",[List.Sum=List.Sum]),

                //-----------------------------------------------------------------------------
                // Résultat avec identification des colonnes titres à gauche des cumuls
                //-----------------------------------------------------------------------------
                ListLeftColumns=Text.Split(pGroupByColumn,";"),
                ToTableFinalGroupBy = Table.Group(Sort, ListLeftColumns, ToEvalTotal)
            in 
                ToTableFinalGroupBy,

        //-----------------------------------------------------------------------------
        // On recupère les binaires des fichiers présents dans le dossier
        // et on construit la table
        //-----------------------------------------------------------------------------
        ToTableFinalGroupBy = if FilterSource is null then null else BuildTable(FilterSource[Content])
        
    in
        //-----------------------------------------------------------------------------
        // Si nous n'obtenons rien car aucune donnée on renvoie une table vide
        //-----------------------------------------------------------------------------        
        if ToTableFinalGroupBy is null then #table({},{}) else ToTableFinalGroupBy                        
in 
    fnCombineMultiTablesEXCELGroupBy_H