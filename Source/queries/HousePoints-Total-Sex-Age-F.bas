﻿dbMemo "SQL" ="SELECT DISTINCT [Report Base].H_NAme, [Report Base].H_Code, [Report Base].H_ID, "
    "[Report Base].Sex, [Report Base].Age, [Report Base].[Sex Sub], [HousePoints-Tota"
    "l-Sex-Age].SumOfPoints\015\012FROM [Report Base] LEFT JOIN [HousePoints-Total-Se"
    "x-Age] ON ([Report Base].H_ID = [HousePoints-Total-Sex-Age].H_ID) AND ([Report B"
    "ase].[Sex Sub] = [HousePoints-Total-Sex-Age].[Sex Sub]) AND ([Report Base].Age ="
    " [HousePoints-Total-Sex-Age].Age);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[Report Base].[Sex Sub]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report Base].H_NAme"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[HousePoints-Total-Sex-Age].SumOfPoints"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report Base].H_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report Base].H_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report Base].Sex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report Base].Age"
        dbLong "AggregateType" ="-1"
    End
End
