Operation =1
Option =8
Begin InputTables
    Name ="House Points-Extra Total"
    Name ="House Points - Total"
End
Begin OutputColumns
    Expression ="[House Points - Total].H_Code"
    Expression ="[House Points - Total].H_NAme"
    Expression ="[House Points - Total].H_ID"
    Expression ="[House Points - Total].SumOfPoints"
    Alias ="Extras"
    Expression ="IIf(IsNull([SumOfNumPts]),0,[SumOfNumPts])"
    Alias ="GrandTotal"
    Expression ="IIf(IsNull([SumOfPoints]),0,[SumOfPoints])+[Extras]"
    Alias ="PercentileTotal"
    Expression ="CalculatePercTotal([GrandTotal],DLookUp(\"[AllPoints]\",\"House Points - All Awa"
        "rded\"),DLookUp(\"[AllPoints]\",\"House Points - All Awarded\"))"
    Expression ="[House Points - Total].CompPool"
    Alias ="%tot"
    Expression ="[GrandTotal]/(SELECT Sum([House Points - Total].[SumOfPoints]) AS TotalPoints FR"
        "OM [House Points - Total])*100"
End
Begin Joins
    LeftTable ="House Points-Extra Total"
    RightTable ="House Points - Total"
    Expression ="[House Points-Extra Total].H_ID = [House Points - Total].H_ID"
    Flag =3
End
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
        dbText "Name" ="[House Points - Total].H_ID"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
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
        dbInteger "ColumnWidth" ="2400"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%tot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total].H_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total].H_NAme"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total].CompPool"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[House Points - Total].SumOfPoints"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =102
    Right =968
    Bottom =688
    Left =-1
    Top =-1
    Right =922
    Bottom =161
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
        Name ="House Points - Total"
        Name =""
    End
End
