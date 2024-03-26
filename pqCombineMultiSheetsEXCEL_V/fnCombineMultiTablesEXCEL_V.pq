let 
    //-----------------------------------------------------------------------------
    // Combinaison d'un dossier de fichiers Excel
    // On ne récupère que les objets tables - tableaux structurés - Kind = Table
    //-----------------------------------------------------------------------------
    fnCombineMultiTablesEXCEL_V = (
            pFolder                    as text,
            optional pFilterFolder     as any,
            optional pFilterExtention  as any,
            optional pFilterFiles      as any,
            optional pFilterTable      as any,
            optional pRefItem          as any
            ) as table =>

        let
            //-----------------------------------------------------------------------------
            // Fonction : On récupère le contenu du dossier
            //-----------------------------------------------------------------------------
            GetSource = () =>
                let 
                    Src = try 
                            Folder.Files(pFolder) 
                          otherwise 
                            #table({},{}),

                    ToSrc = try 
                                Src 
                            otherwise 
                                #table({},{}),

                    //------------------------------
                    // Filtre sur les dossiers
                    //-------------------------------
                    TbFilterFolder =  try if pFilterFolder is null or pFilterFolder = false then 
                                            ToSrc 
                                            else Table.SelectRows(ToSrc, each [Folder Path] = pFilterFolder)
                                      otherwise 
                                            #table({},{}),

                    //------------------------------
                    // Filtre sur l'extension
                    //------------------------------
                    TbFilterExtension = try if pFilterExtention is null or pFilterExtention = false then 
                                            TbFilterFolder 
                                            else Table.SelectRows(TbFilterFolder, each Text.Contains([Extension], pFilterExtention))
                                        otherwise
                                            #table({},{}),

                    //------------------------------
                    // Filtre sur les fichiers
                    //------------------------------
                    TbFilterFiles =  try if pFilterFiles is null or pFilterFiles = false then 
                                            TbFilterExtension
                                            else Table.SelectRows(TbFilterExtension, each Text.StartsWith([Name], pFilterFiles))
                                     otherwise
                                            #table({},{})
                in 
                    TbFilterFiles,

            //-----------------------------------------------------------------------------
            // Fonction : On récupère la liste des binaires
            //-----------------------------------------------------------------------------
            GetContent = () =>
                let 
                    SrcContent = Source[Content],
                    ToContent = try 
                                    SrcContent 
                                otherwise 
                                    #table({},{})
                in 
                    SrcContent,

            //-----------------------------------------------------------------------------
            // Fonction : Construction de la table finale
            //-----------------------------------------------------------------------------
            Build = () =>
            let 
                //-----------------------------------------------------------------------------
                // Pour chaque binaire on récupère les tables (tableaux structurés)
                //-----------------------------------------------------------------------------
                TbListCombine = List.Accumulate(
                                    Content,
                                    #table({},{}),                        
                                    (state,current) =>  
                                            let 
                                                Classeur=Excel.Workbook(current),
                                                // On ne retient que les tableaux structurés
                                                Filter=Table.SelectRows(Classeur, each ([Kind] = "Table")),
                                                CombineFilter=Table.Combine({state,Filter})
                                            in 
                                                CombineFilter                               
                                ),
                                
                //-----------------------------------------------------------------------------
                // Filtre éventuel sur les noms des tables à récupérer 
                //-----------------------------------------------------------------------------
                FilterTable = try if pFilterTable is null or pFilterTable = false then 
                                     TbListCombine 
                                     else Table.SelectRows(TbListCombine, each Text.StartsWith([Item], pFilterTable))
                              otherwise 
                                     #table({},{}),   
                
                //-----------------------------------------------------------------------------
                // On extrait le les colonnes du 1er classeur qui va servir de référence 
                //-----------------------------------------------------------------------------
                ListData=if Table.IsEmpty(FilterTable) then 
                            {} 
                            else FilterTable[Data],

                ColNameRef={"Item"},

                ListColumns= if List.IsEmpty(ListData) then 
                                {} else 
                                Table.ColumnNames(ListData{0}),
                
                //-----------------------------------------------------------------------------
                // On récupère les data dans les binaires pour les convertir en table
                //-----------------------------------------------------------------------------
                ToData = Table.ExpandTableColumn(FilterTable, "Data", ListColumns),

                //---------------------------------------------------------------------------------
                // On ne retient que les colonnes de référence avec Item (nom de la table) ou pas
                //---------------------------------------------------------------------------------
                ToSelectColumns = if pRefItem is null or pRefItem = false then 
                                     Table.SelectColumns(ToData, ListColumns)
                                     else Table.SelectColumns(ToData, ColNameRef & ListColumns)
            in
                ToSelectColumns,

            //-----------------------------------------------------------------------------
            // MAIN PROCEDURE
            //-----------------------------------------------------------------------------
            Source = GetSource(),

            Content = if not Table.IsEmpty(Source) then 
                        GetContent() 
                        else #table({"ERROR"},{{"#ERROR EMPTY"}}),

            ToBuild = if List.IsEmpty(Content) then 
                        #table({},{}) 
                        else Build()
        in 
            ToBuild
in 
    fnCombineMultiTablesEXCEL_V