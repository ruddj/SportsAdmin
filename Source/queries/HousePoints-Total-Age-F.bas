﻿Operation =1
Option =2
Begin InputTables
    Name ="HousePoints-Total-Age"
    Name ="Report Base"
End
Begin OutputColumns
    Expression ="[Report Base].Age"
    Expression ="[HousePoints-Total-Age].SumOfPoints"
    Expression ="[Report Base].H_NAme"
    Expression ="[Report Base].H_Code"
    Expression ="[Report Base].H_ID"
End
Begin Joins
    LeftTable ="Report Base"
    RightTable ="HousePoints-Total-Age"
    Expression ="[Report Base].H_ID = [HousePoints-Total-Age].H_ID"
    Flag =2
    LeftTable ="Report Base"
    RightTable ="HousePoints-Total-Age"
    Expression ="[Report Base].Age = [HousePoints-Total-Age].Age"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =927
    Bottom =710
    Left =-1
    Top =-1
    Right =909
    Bottom =207
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =14
        Top =6
        Right =134
        Bottom =184
        Top =0
        Name ="HousePoints-Total-Age"
        Name =""
    End
    Begin
        Left =213
        Top =5
        Right =392
        Bottom =153
        Top =0
        Name ="Report Base"
        Name =""
    End
End
