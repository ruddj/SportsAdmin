Operation =1
Option =8
Begin InputTables
    Name ="House Points-Extra Total"
    Name ="House Points - Total (All Events)"
End
Begin OutputColumns
    Expression ="[House Points - Total (All Events)].H_Code"
    Expression ="[House Points - Total (All Events)].H_NAme"
    Expression ="[House Points - Total (All Events)].H_ID"
    Expression ="[House Points - Total (All Events)].SumOfPoints"
    Alias ="Extras"
    Expression ="IIf(IsNull([SumOfNumPts]),0,[SumOfNumPts])"
    Alias ="GrandTotal"
    Expression ="IIf(IsNull([SumOfPoints]),0,[SumOfPoints])+[Extras]"
    Alias ="PercentileTotal"
    Expression ="CalculatePercTotal([GrandTotal],DLookUp(\"[AllPoints]\",\"House Points - All Awa"
        "rded\"),DLookUp(\"[AllPoints]\",\"House Points - All Awarded\"))"
    Expression ="[House Points - Total (All Events)].CompPool"
    Alias ="%tot"
    Expression ="[GrandTotal]/[CompPool]*100"
End
Begin Joins
    LeftTable ="House Points-Extra Total"
    RightTable ="House Points - Total (All Events)"
    Expression ="[House Points-Extra Total].H_ID = [House Points - Total (All Events)].H_ID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Extras"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PercentileTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%tot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total (All Events)].H_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total (All Events)].H_NAme"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total (All Events)].CompPool"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total (All Events)].H_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total (All Events)].SumOfPoints"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =46
    Top =183
    Right =1008
    Bottom =635
    Left =-1
    Top =-1
    Right =944
    Bottom =177
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="House Points-Extra Total"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="House Points - Total (All Events)"
        Name =""
    End
End
