let
    //----------------------------------------------------------------------------------------------------
    // fnSelectCase
    //----------------------------------------------------------------------------------------------------

    //----------------------------------------------------------------------------------------------------
    // Cette fonction simule le SELECT CASE en VBA ou le CHOISIR en formule
    // ou le SWITCH en DAX
    //----------------------------------------------------------------------------------------------------
    // Elle renvoie un type :
    // - text
    // - number
    // - date
    // - table
    // - list
    // - list(list)
    // - list(table)
    // en cas d'incohérences entre la liste pSelect et pCase ou erreur d'évaluation
    // c'est pValue qui est renvoyé afin de ne pas planté ou vide selon les cas
    //------------------------------------------------------------------------------------------------------
    SelectCase = (
            pValue              as any,  // valeur recherchée
            pSelect             as any,  // équivalent du SELECT CASE valeur
            pCase               as any,  // les différents cas d'évaluation comme le CASE condition + retour en VBA
            optional pDefault   as any,  // si la valeur n'est pas trouvée
            optional pSep       as any,  // séparateur quand la liste est passée sous forme de chaine
            optional pType      as any   // type renvoyé          
        ) as any =>                      // étant que le type renvoyé peut être de plusieurs formes

    let
        //----------------------------------------------------------------------------------------------
        // Pa défaut le séparateur est ;
        //----------------------------------------------------------------------------------------------
        Separator=if pSep is null then ";" else pSep,

        //----------------------------------------------------------------------------------------------
        // Sécurisation si la liste de recherche de la clef est vide
        // pList1 : liste contenant la liste des éléments à rechercher
        //----------------------------------------------------------------------------------------------
        pList1= if pSelect is null then null else if pSelect is list
                  then
                       if List.IsEmpty(pSelect) then {pValue} else List.Buffer(pSelect)
                  else if pSelect is null then
                          pValue
                          // Conversion en liste dans un buffer
                          else List.Buffer(Text.Split(Text.Trim(pSelect),Separator)),

        //----------------------------------------------------------------------------------------------
        // Sécurisation si la liste de valeurs à retourner est vide
        // pList2 : liste contenant les valeurs de retour en fonction de la position dans la liste 1
        //----------------------------------------------------------------------------------------------
        pList2=if pCase is null then null else if pCase is list
                  then
                       if List.IsEmpty(pCase) then {pValue} else List.Buffer(pCase)
                  else if pCase is null then
                          pValue
                          // Conversion en liste dans un buffer
                          else List.Buffer(Text.Split(Text.Trim(pCase),Separator)),      

        //----------------------------------------------------------------------------------------------
        // Sécurisation pour ne pas planter la recherche dans la liste 1
        // recherche de la clef pValue dans la liste 1
        // on renvoie la position (index) dans la liste 1
        //----------------------------------------------------------------------------------------------
        Position=if pList1 is null then
                    null
                    else List.PositionOf (pList1,pValue),

        //----------------------------------------------------------------------------------------------
        // Sécurisation pour ne pas planter l'extraction de la liste 2
        // on va recherche l'élement correspondant dans la liste 2 des valeurs de retour
        //----------------------------------------------------------------------------------------------
        ReturnValue=if pList2 is null or Position is null then
                       null
                       else if Position=-1 then
                               pDefault
                               //-----------------------------------------------------------------------
                               // Sécurisation si pas le même nombre d'éléments entre les listes 1 et 2
                               //-----------------------------------------------------------------------
                               else if Position > List.Count(pList2) -1 then null else pList2{Position},

        //----------------------------------------------------------------------------------------------
        // Evalutation  du type de retour
        // si aucune valeur n'est possible on preend le parti de renvoyer la valeur à évaluer (pValue)
        // Ainsi on ne plante pas la fonction
        //----------------------------------------------------------------------------------------------
        EvalReturn=if ReturnValue is null then
                      pValue
                      else if pType = "any" or pType is null then
                                ReturnValue
                                else if pType = "table" then
                                        Expression.Evaluate(ReturnValue,#shared)  // nécessaire pour traduire la table
                                        else if pType = "list" then
                                                {ReturnValue}
                                                else if pType = "number" then
                                                        Number.FromText(ReturnValue)
                                                        else if pType = "date" then
                                                                Date.FromText(ReturnValue)
                                                                else ReturnValue

    in
       EvalReturn

in
    SelectCase