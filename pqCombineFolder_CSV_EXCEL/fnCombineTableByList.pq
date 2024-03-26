let
    //---------------------------------------------------------------------------------
    // Fonction permettant de combiner plusieurs tables (tableaux structurés)
    // Ce dossier et ces tables sont définis dans une table paramètres
    //---------------------------------------------------------------------------------

    // La table des paramètres et le paramètre tables à combiner
    fnCombineExcelByList = (pTableParams as text, pParam as text) as table =>
    
    let
        LstTables=Text.Split(fnGetParameter(pTableParams,pParam),";"),   // Liste des tables à combiner
        LstCombine=List.Accumulate(
                        LstTables,            // Initialisation de la boucle sur la liste à traiter 
                        #table({},{}),       // Initialisation de l'accumulateur sur une table vide 
                        (state,current) =>   // Boucle de traitement pour combiner les fichiers un à un
                            let 
                                Source = fnGetTable(current),     // Importation de la table en cours
                                Combine = Table.Combine( {state , Source})  // Combinaison de la table
                            in                            
                            Combine   // Ensemble des tables combinées
                        )      
    in
        LstCombine
in 
    fnCombineExcelByList