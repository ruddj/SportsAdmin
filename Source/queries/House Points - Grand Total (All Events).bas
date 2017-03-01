dbMemo "SQL" ="SELECT DISTINCTROW [House Points - Total (All Events)].H_Code, [House Points - T"
    "otal (All Events)].H_NAme, [House Points - Total (All Events)].H_ID, [House Poin"
    "ts - Total (All Events)].SumOfPoints, IIf(IsNull([SumOfNumPts]),0,[SumOfNumPts])"
    " AS Extras, IIf(IsNull([SumOfPoints]),0,[SumOfPoints])+[Extras] AS GrandTotal, C"
    "alculatePercTotal([GrandTotal],[House Points - Total (All Events)].[H_ID],[CompP"
    "ool]) AS PercentileTotal, [House Points - Total (All Events)].CompPool, [GrandTo"
    "tal]/[CompPool]*100 AS [%tot]\015\012FROM [House Points-Extra Total] RIGHT JOIN "
    "[House Points - Total (All Events)] ON [House Points-Extra Total].H_ID = [House "
    "Points - Total (All Events)].H_ID;\015\012"
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
        dbText "Name" ="Extras"
    End
    Begin
        dbText "Name" ="GrandTotal"
    End
    Begin
        dbText "Name" ="PercentileTotal"
    End
    Begin
        dbText "Name" ="%tot"
    End
End
