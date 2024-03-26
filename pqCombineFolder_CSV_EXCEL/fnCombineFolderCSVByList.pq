let
    //---------------------------------------------------------------------------------
    // Fonction permettant de combiner plusieurs fichiers spécifiques dans un dossier
    // Ce dossier et ces fichiers sont définis dans une table paramètres
    //---------------------------------------------------------------------------------

    // La table des paramètres et le paramètre contenu les fichiers à combiner
    fnCombineByList = (pTableParams as text, pParam as text) as table =>
    let
        LstFiles=Text.Split(fnGetParameter(pTableParams,pParam),";"),   // Liste des fichiers à combiner
        LstCombine=List.Accumulate(
                        LstFiles,            // Initialisation de la boucle sur la liste à traiter 
                        #table({},{}),       // Initialisation de l'accumulateur sur une table vide 
                        (state,current) =>   // Boucle de traitement pour combiner les fichiers un à un
                            let 
                                Source = fnGetCSV(current),     // Importation du fichier en cours
                                Combine = Table.Combine( {state , Source})  // Combinaison du fichier en cours avec ceux déja combinés dans state
                            in                            
                            Combine   // Ensemble des fichiers combinés
                        )      
    in
        LstCombine
in 
    fnCombineByList    // Résultat de la fonction