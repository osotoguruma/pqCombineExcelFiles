let fnGetCSV = (pFile as any) as table =>

    //---------------------------------------------------------------------------------
    // Imporatation d'un fichier à partir d'un dossier
    //---------------------------------------------------------------------------------

    let

        //---------------------------------------------------------------------------------
        // Lecture des paramètres
        //---------------------------------------------------------------------------------
        pFolder=fnGetParameter("TB_PARAMS","CHEMIN_CSV"),
        pNbCols=fnGetParameter("TB_PARAMS","NB_COLS"),
        pDelim=fnGetParameter("TB_PARAMS","DELIMITEUR"),    
        pEncoding=fnGetParameter("TB_PARAMS","ENCODE"),

        //---------------------------------------------------------------------------------
        // Paramètres techniques d'importation d'un fichier .csv
        //---------------------------------------------------------------------------------
        pParamCSV=[Delimiter=pDelim, Columns=pNbCols, Encoding=pEncoding, QuoteStyle=QuoteStyle.None],

        //---------------------------------------------------------------------------------
        // Le combine nécessite une liste de table à combiner
        //---------------------------------------------------------------------------------
        pPath=pFolder & "\" & pFile,

        //---------------------------------------------------------------------------------
        // Importation du fichier
        //---------------------------------------------------------------------------------
        Source = Csv.Document(File.Contents(pPath),pParamCSV),

        //---------------------------------------------------------------------------------
        // Titre des colonnes
        //---------------------------------------------------------------------------------
        ToTable = Table.PromoteHeaders(Source, [PromoteAllScalars=true])

    in
        ToTable
in 
    try fnGetCSV otherwise null