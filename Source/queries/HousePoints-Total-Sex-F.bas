﻿Operation =1
Option =2
Begin InputTables
    Name ="Report Base"
    Name ="HousePoints-Total-Sex"
End
Begin OutputColumns
    Expression ="[Report Base].H_NAme"
    Expression ="[Report Base].H_Code"
    Expression ="[Report Base].H_ID"
    Expression ="[HousePoints-Total-Sex].SumOfPoints"
    Expression ="[HousePoints-Total-Sex].[Sex Sub]"
End
Begin Joins
    LeftTable ="Report Base"
    RightTable ="HousePoints-Total-Sex"
    Expression ="[Report Base].H_ID = [HousePoints-Total-Sex].H_ID"
    Flag =2
    LeftTable ="Report Base"
    RightTable ="HousePoints-Total-Sex"
    Expression ="[Report Base].[Sex Sub] = [HousePoints-Total-Sex].[Sex Sub]"
    Flag =2
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
        dbText "Name" ="[HousePoints-Total-Sex].[Sex Sub]"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =143
    Top =142
    Right =1658
    Bottom =455
    Left =-1
    Top =-1
    Right =1497
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Report Base"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="HousePoints-Total-Sex"
        Name =""
    End
End
