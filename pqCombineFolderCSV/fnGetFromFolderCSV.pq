let fnGetFromFolder = (pPath as text, pDelim as text, pEncoding as number) as table =>

    //---------------------------------------------------------------------------------
    // Combinaison du contenu d'un dossier
    // Ici nous passons pas par le code mis automatiquement quand on fait appel à l'assistant
    // Limportation s'effectue directement via les binaires des fichiers
    //---------------------------------------------------------------------------------

    let
        SOURCE_CONTENT="Content",

        //-------------------------------------------------------------------------------------
        // Fonction pour la lecture des binaires bufferisés
        //-------------------------------------------------------------------------------------
        fnReadBinaryFile=(pFile as binary) as table =>         
            Table.PromoteHeaders( Csv.Document(pFile, [Delimiter=pDelim,Encoding=pEncoding])),
        
        //-------------------------------------------------------------------------------------
        // Définition de la source de données
        //-------------------------------------------------------------------------------------
        Source = Folder.Files(pPath),

        //-------------------------------------------------------------------------------------
        // Bufferisation des binaires
        //-------------------------------------------------------------------------------------
        DrillDownContent = Table.SelectColumns(Source,{SOURCE_CONTENT}),
        ListTransform = List.Transform (DrillDownContent[Content], Binary.Buffer),

        //-------------------------------------------------------------------------------------
        // Transformation et lecture des binaires
        //-------------------------------------------------------------------------------------
        ReadBinaryFiles = List.Transform(ListTransform, fnReadBinaryFile),

        //-------------------------------------------------------------------------------------
        // Combinaison du fichier
        //-------------------------------------------------------------------------------------
        TableCombine = Table.Combine(ReadBinaryFiles )
    in
        TableCombine
in 
    try fnGetFromFolder otherwise null