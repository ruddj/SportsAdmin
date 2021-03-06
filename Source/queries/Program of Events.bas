﻿Operation =1
Option =0
Where ="(((EventType.Flag)=Yes) AND ((EventType.Include)=Yes) AND ((Events.Include)=Yes)"
    " AND ((Events.Sex) Like DLookUp(\"[Rsex]\",\"Misc-EventLists\")) AND ((Events.Ag"
    "e) Like DLookUp(\"[Rage]\",\"Misc-EventLists\")) AND ((Heats.Heat) Like DLookUp("
    "\"[Rheat]\",\"Misc-EventLists\")) AND ((Heats.F_Lev) Like DLookUp(\"[Rfinal]\",\""
    "Misc-EventLists\")))"
Begin InputTables
    Name ="EventType"
    Name ="Events"
    Name ="Heats"
    Name ="Sex Sub"
    Name ="Final Level Sub"
    Name ="House"
End
Begin OutputColumns
    Expression ="House.H_Code"
    Expression ="EventType.Flag"
    Expression ="EventType.Include"
    Expression ="Events.Include"
    Expression ="Heats.E_Code"
    Expression ="Heats.E_Number"
    Expression ="EventType.ET_Des"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="Events.Record"
    Expression ="Events.RecName"
    Expression ="Events.RecHouse"
    Expression ="Heats.Heat"
    Expression ="[Final Level Sub].F_Lev_Sub"
    Expression ="Heats.F_Lev"
    Expression ="Heats.Status"
    Alias ="E_Time2"
    Expression ="IIf([E_Time]>=1,Format([E_Time],\"d-mmm h:nn am/pm\"),Format([E_Time],\"h:nn am/"
        "pm\"))"
    Expression ="EventType.Units"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="Sex Sub"
    Expression ="Events.Sex = [Sex Sub].Sex"
    Flag =1
    LeftTable ="Events"
    RightTable ="House"
    Expression ="Events.RecHouse = House.H_ID"
    Flag =2
    LeftTable ="Events"
    RightTable ="Heats"
    Expression ="Events.E_Code = Heats.E_Code"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =1
    LeftTable ="Final Level Sub"
    RightTable ="Heats"
    Expression ="[Final Level Sub].F_Lev = Heats.F_Lev"
    Flag =1
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
        dbText "Name" ="E_Time2"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
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
    Bottom =344
    Left =0
    Top =12
    ColumnsShown =539
    Begin
        Left =38
        Top =-6
        Right =134
        Bottom =176
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =172
        Top =-6
        Right =268
        Bottom =176
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =302
        Top =6
        Right =398
        Bottom =203
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =28
        Top =191
        Right =124
        Bottom =268
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =143
        Top =206
        Right =239
        Bottom =283
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =436
        Top =6
        Right =532
        Bottom =202
        Top =0
        Name ="House"
        Name =""
    End
End
