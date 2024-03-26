let
    //---------------------------------------------------------------------------------
    // Fonction permettant de combiner plusieurs fichiers spécifiques dans un dossier
    // Ce dossier et ces fichiers sont définis dans une table paramètres
    // Soit EXCEL
    // Soit CSV
    //---------------------------------------------------------------------------------

    // La table des paramètres et le paramètre contenu les fichiers à combiner
    fnCombineExcelByList = (pTableParams as text, pParam as text, optional ppType as text, optional ppItem as text, optional ppKind as text) as table =>
    
    let
        //-------------------------------------------------------------------
        // Permet de gérer soit un import CSV soit un import EXCEL
        //-------------------------------------------------------------------
        pType=if ppType is null then "CSV" else ppType,
        LstFiles=Text.Split(fnGetParameter(pTableParams,pParam),";"),   // Liste des fichiers à combiner
        LstCombine=List.Accumulate(
                        LstFiles,            // Initialisation de la boucle sur la liste à traiter 
                        #table({},{}),       // Initialisation de l'accumulateur sur une table vide 
                        (state,current) =>   // Boucle de traitement pour combiner les fichiers un à un
                            let 
                                //------------------------------------------------------
                                // Aiguillage soit EXCEL soit CSV
                                //------------------------------------------------------
                                Source = if Text.Upper(pType)="EXCEL" then 
                                         fnGetEXCEL(current, ppItem, ppKind)     // On va chercher un fichier Excel
                                         else if Text.Upper(pType)="CSV" then 
                                                 fnGetCSV(current)               // On va chercher un fichier CSV
                                                 else null,                      // Sinon on prend le parti de ne rien renvoyer

                                Combine = Table.Combine( {state , Source})       // Combinaison du fichier en cours avec ceux déja combinés dans state
                            in                            
                            Combine   // Ensemble des fichiers combinés
                        )      
    in
        LstCombine
in 
    fnCombineExcelByList