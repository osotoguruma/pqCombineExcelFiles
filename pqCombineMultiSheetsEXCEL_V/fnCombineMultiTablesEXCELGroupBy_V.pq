let 
    //-----------------------------------------------------------------------------
    // Apr�s avoir fusionner les fichiers Excel du dossier
    // On peut faire un regroupement sur la colonne ma�tre de pivotement
    //-----------------------------------------------------------------------------

    fnGroupByV = (
            pFolder                    as text,
            pFieldGroupBy              as text,
            pFieldColValue             as text,
            pFieldValueToSum           as text,
            optional pFilterFolder     as any,
            optional pFilterExtention  as any,
            optional pFilterFiles      as any,
            optional pFilterTable      as any,
            optional pRefItem          as any
    ) as table =>

        //-----------------------------------------------------------------------------
        // Combinaison du dossier Excel
        //-----------------------------------------------------------------------------
        let
            Source = fnCombineMultiTablesEXCEL_V(
                pFolder,
                pFilterFolder,
                pFilterExtention,
                pFilterFiles,
                pFilterTable,
                pRefItem
                ),
                
            //-----------------------------------------------------------------------
            // Evaluation de la colonne pivot pour effectuer la somme
            //-----------------------------------------------------------------------
            StrEval="each List.Sum([" & pFieldValueToSum& "])",
            EvalListSum=Expression.Evaluate(StrEval,[List.Sum=List.Sum]),

            //-----------------------------------------------------------------------
            // Regroupement de la colonne pivot
            //-----------------------------------------------------------------------
            ToTableGroupBy = Table.Group(Source, {pFieldGroupBy}, {{pFieldColValue,  EvalListSum, type number}})

        in
            ToTableGroupBy
in 
    fnGroupByV