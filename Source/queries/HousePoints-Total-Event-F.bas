﻿Operation =1
Option =2
Begin InputTables
    Name ="Report Base"
    Name ="HousePoints-Total-Event"
End
Begin OutputColumns
    Expression ="[Report Base].H_NAme"
    Expression ="[Report Base].H_Code"
    Expression ="[Report Base].H_ID"
    Expression ="[HousePoints-Total-Event].SumOfPoints"
    Expression ="[Report Base].ET_Des"
End
Begin Joins
    LeftTable ="Report Base"
    RightTable ="HousePoints-Total-Event"
    Expression ="[Report Base].H_ID = [HousePoints-Total-Event].H_ID"
    Flag =2
    LeftTable ="Report Base"
    RightTable ="HousePoints-Total-Event"
    Expression ="[Report Base].ET_Des = [HousePoints-Total-Event].ET_Des"
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
    Left =137
    Top =136
    Right =1002
    Bottom =441
    Left =-1
    Top =-1
    Right =847
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =26
        Top =6
        Right =152
        Bottom =113
        Top =0
        Name ="Report Base"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="HousePoints-Total-Event"
        Name =""
    End
End
