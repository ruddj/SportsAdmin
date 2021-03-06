﻿Operation =1
Option =8
Begin InputTables
    Name ="CompEvents"
    Name ="Competitors"
    Name ="Heats"
    Name ="EventType"
    Name ="Events"
End
Begin OutputColumns
    Expression ="Competitors.PIN"
    Expression ="EventType.ET_Des"
    Expression ="CompEvents.Heat"
    Expression ="CompEvents.Place"
    Alias ="res"
    Expression ="IIf(IsNull([Result]),\"-\",[Result]) & \" \" & [units]"
    Expression ="CompEvents.Points"
    Expression ="CompEvents.F_Lev"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="CompEvents"
    Expression ="Events.E_Code = CompEvents.E_Code"
    Flag =3
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =3
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
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
        dbText "Name" ="res"
    End
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
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Events"
        Name =""
    End
End
