﻿Operation =1
Option =0
Where ="(((House.Flag)=True) AND ((House.Include)=True))"
Begin InputTables
    Name ="CompEvents"
    Name ="Competitors"
    Name ="House"
    Name ="Heats"
    Name ="Events"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="House.H_ID"
    Alias ="E_Num"
    Expression ="IIf(IsNull([E_Number]),0,[E_Number])"
    Expression ="CompEvents.Points"
End
Begin Joins
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =1
    LeftTable ="Events"
    RightTable ="Heats"
    Expression ="Events.E_Code = Heats.E_Code"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
    Flag =1
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =1
End
Begin OrderBy
    Expression ="House.H_ID"
    Flag =0
    Expression ="IIf(IsNull([E_Number]),0,[E_Number])"
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
    Begin
        dbText "Name" ="E_Num"
        dbInteger "ColumnWidth" ="1080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =3
    Right =770
    Bottom =405
    Left =-1
    Top =-1
    Right =676
    Bottom =203
    Left =0
    Top =27
    ColumnsShown =539
    Begin
        Left =336
        Top =-13
        Right =432
        Bottom =94
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =452
        Top =-17
        Right =548
        Bottom =90
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =583
        Top =-24
        Right =679
        Bottom =83
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =216
        Top =-12
        Right =312
        Bottom =95
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =118
        Top =15
        Right =214
        Bottom =122
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =10
        Top =31
        Right =106
        Bottom =138
        Top =0
        Name ="EventType"
        Name =""
    End
End
