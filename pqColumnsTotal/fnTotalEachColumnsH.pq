let

   fnTotalEachColumnsH =(pTable as any, pOffset as number, pLibTotalH as text, optional pBuff as logical) as table =>

   let

      //-----------------------------------------------------------------------------------
      // Fonction permettant de totaliser les colonnes d'un tableau
      //-----------------------------------------------------------------------------------

      /*
      pTable="TB_VENTES",
      pOffset=1,
      pLibTotalH="Total"
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
      // Nombre de colonnes à la source
      //-----------------------------------------------------------------------------------
      ColNamesSource = List.Buffer(Table.ColumnNames(bSource)),
      FirstN=List.FirstN(ColNamesSource,pOffset),

      //-----------------------------------------------------------------------------------
      // On ajoute la colonne total en fin de ligne
      //-----------------------------------------------------------------------------------
      TblAddColumnTotal = Table.AddColumn(
                          bSource,
                          pLibTotalH,  
                          each List.Sum(   // Pour chaque ligne on fait la somme
                                    Record.ToList(  // On prend le parti de faire somme de toutes les colonnes
                                           Record.SelectFields(_,   // mais on supprime les colonnes de tête
                                           List.RemoveItems(Table.ColumnNames(Source), FirstN)))))

   in
      TblAddColumnTotal
in 
   fnTotalEachColumnsH