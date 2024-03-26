let
    //-----------------------------------------------------------------------
    // Après combinaison du dossier - pivot de la colonne de regroupement
    //-----------------------------------------------------------------------

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
            // Pivot du tableau
            //-----------------------------------------------------------------------

            //-----------------------------------------------------------------------
            // On détermine la liste des colonnes
            //-----------------------------------------------------------------------

            ListColumns= if pRefItem is null or pRefItem = false then 
                            Table.ColumnNames(Source)
                            else List.LastN(Table.ColumnNames(Source),List.Count(Table.ColumnNames(Source)-1)),            

            //-----------------------------------------------------------------------
            // Evaluation de la colonne pivot
            //-----------------------------------------------------------------------
            StrEval="List.Distinct(Source[" & pFieldGroupBy & "])" ,
            EvalListDistinct=Expression.Evaluate(StrEval,[List.Distinct=List.Distinct, Source=Source]),

            //-----------------------------------------------------------------------
            // Pivotement de la colonne maître
            //-----------------------------------------------------------------------
            ToTablePivot = Table.Pivot(Source, EvalListDistinct, pFieldGroupBy, pFieldValueToSum, List.Sum)

        in
            ToTablePivot
in 
    fnGroupByV