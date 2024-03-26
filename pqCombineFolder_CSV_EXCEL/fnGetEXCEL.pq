let fnGetEXCEL = (pFolder as text, pFile as text, optional ppItem as text, optional ppKind as text) as table =>

    //---------------------------------------------------------------------------------
    // Imporatation d'un fichier Excel à partir d'un dossier
    //---------------------------------------------------------------------------------

    let

        //---------------------------------------------------------------------------------
        // Lecture des paramètres
        //---------------------------------------------------------------------------------
        pKind=if ppKind is null then "Sheet" else ppKind,
        pItem=if ppItem is null then "Feuil1" else ppItem,
        

        //---------------------------------------------------------------------------------
        // Le combine nécessite une liste de table à combiner
        //---------------------------------------------------------------------------------
        pPath=pFolder & "\" & pFile,

        //---------------------------------------------------------------------------------
        // Importation du fichier
        //---------------------------------------------------------------------------------
        Source = Excel.Workbook(File.Contents(pPath), null, true),
        ToTable=Source{[Item=pItem,Kind=Text.Proper(pKind)]}[Data],
        Promote=if Text.Upper(pKind)="SHEET" then 
                   Table.PromoteHeaders(ToTable, [PromoteAllScalars=true])
                   else if Text.Upper(pKind)="TABLE" then 
                        ToTable
                        else Source

    in
        Promote
in 
    try fnGetEXCEL otherwise null