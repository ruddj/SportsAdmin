﻿Operation =1
Option =8
Begin InputTables
    Name ="Heats"
    Name ="Events"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="Heats.F_Lev"
    Expression ="EventType.ET_Des"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="Heats.Heat"
    Expression ="Heats.E_Number"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="Heats"
    Expression ="Events.E_Code = Heats.E_Code"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =1
End
Begin OrderBy
    Expression ="Heats.F_Lev"
    Flag =1
    Expression ="EventType.ET_Des"
    Flag =0
    Expression ="Events.Sex"
    Flag =0
    Expression ="Events.Age"
    Flag =0
    Expression ="Heats.Heat"
    Flag =0
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
    Bottom =393
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="EventType"
        Name =""
    End
End
