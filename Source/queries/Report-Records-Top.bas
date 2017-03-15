dbMemo "SQL" ="SELECT [Report-Records].ET_Des, [Report-Records].Sex, [Report-Records].Age, [Rep"
    "ort-Records].BestResult, [Report-Records].ET_Code, [Report-Records].E_Code, [Rep"
    "ort-Records].Units, [Report-Records].AgeSex, [Report-Records].Result, [Report-Re"
    "cords].ResultFormated, [Report-Records].FullName, [Report-Records].CompetitorHou"
    "se, [Report-Records].[Sex Sub], [Report-Records].Flag, [Report-Records].Date\015"
    "\012FROM [Report-Records]\015\012WHERE [Report-Records].BestResult IN \015\012(S"
    "ELECT TOP 1 TopRec.BestResult\015\012FROM [Report-Records] AS TopRec\015\012WHER"
    "E TopRec.E_Code = [Report-Records].E_Code\015\012ORDER BY TopRec.BestResult);\015"
    "\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="[Report-Records].Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].ET_Des"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].ET_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].Sex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].AgeSex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].Age"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].BestResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].E_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].Result"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].ResultFormated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].CompetitorHouse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].[Sex Sub]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Report-Records].Flag"
        dbLong "AggregateType" ="-1"
    End
End
