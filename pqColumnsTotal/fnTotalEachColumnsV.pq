let

   fnTotalEachColumnsV =(pTable as any, pOffset as number, pLibTotal as text, optional pBuff as logical) as table =>

   let

      //-----------------------------------------------------------------------------------
      // Fonction permettant de totaliser les colonnes d'un tableau
      //-----------------------------------------------------------------------------------

      /*
      pTable="TB_VENTES",
      pOffset=1,
      pLibTotal="Total ventes",
      pBuff=true
      */

      //-----------------------------------------------------------------------------------
      // On importe la source en passant par les buffers pour optimiser les calculs
      // Fait-on le choix de bufferiser la table des données
      //-----------------------------------------------------------------------------------
      Source = if pTable is table then 
                  pTable
                  else Excel.CurrentWorkbook(){[Name = pTable]}[Content],

      // Choix si buffer
      bSource=if pBuff is null then 
                 Source
                 else if pBuff then 
                 Table.Buffer(Source)
                 else Source ,

      //-----------------------------------------------------------------------------------
      // On récupère les colonnes de la table des données
      // Là on peut bufferiser sans risque sur le nbr de colonnes
      //-----------------------------------------------------------------------------------
      ColNames = List.Buffer(Table.ColumnNames(bSource)),

      //-----------------------------------------------------------------------------------
      // On ne retient que les colonnes à cumuler en écartant celles de gauche
      //-----------------------------------------------------------------------------------
      TotalsHeaders=List.Buffer(List.LastN(ColNames,List.Count(ColNames) - pOffset)),
      
      //-----------------------------------------------------------------------------------
      // On crée une table temporaire qui va contenir le total de chaque colonne
      // Et conserve le titre de la 1ère colonne par ColNames{0}
      //-----------------------------------------------------------------------------------
      TblFirstColTmp = Table.FromColumns({{pLibTotal}}, {ColNames{0}}),
      
      //-----------------------------------------------------------------------------------
      // On fait le cumul de chaque colonne dans la table temporaire
      // Boucle pour les cumuls par mois
      // Une manière de mettre en oeuvre List.Accumulate
      //-----------------------------------------------------------------------------------
      CumulColumnsTblTmp = List.Accumulate(
         TotalsHeaders,        // Liste des colonnes à cumuler
         TblFirstColTmp,       // Initialisation sur la 1ère colonne de la table temporaire
         (state, current) =>   // On va boucler sur les colonnes à totaliser 
                               // La colonne en cours de traitement est totalisée
                               // Table.Column(bSource, current) ==> représente la colonne en cours
                               // traduite sous forme de liste pour faire le total de cette liste
         Table.AddColumn(state, current, each List.Sum(Table.Column(bSource, current)), type number)
      ),

      //-----------------------------------------------------------------------------------
      // On fusion la table des cumuls à la table des données pour le tableau final
      // Petite astuce pour ajouter un enregistrement à la fin d'un tableau
      //-----------------------------------------------------------------------------------
      ToTable = Table.Combine({bSource, CumulColumnsTblTmp})

   in
      ToTable
in 
   fnTotalEachColumnsV