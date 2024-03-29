﻿Operation =1
Option =8
Begin InputTables
    Name ="EventType"
    Name ="Events"
    Name ="CompEvents"
    Name ="Competitors"
    Name ="House"
End
Begin OutputColumns
    Expression ="Events.ET_Code"
    Expression ="EventType.ET_Des"
    Expression ="Events.Age"
    Expression ="Events.Sex"
    Alias ="FullName"
    Expression ="RTrim([Surname])+\", \"+RTrim([Gname])"
    Expression ="House.H_NAme"
    Expression ="CompEvents.Heat"
    Expression ="Competitors.PIN"
    Alias ="A_Lane"
    Expression ="IIf([CompEvents].[Lane]+1=1,[House].[Lane],[CompEvents].[Lane])"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="CompEvents"
    Expression ="Events.E_Code = CompEvents.E_Code"
    Flag =1
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =1
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =1
End
Begin OrderBy
    Expression ="EventType.ET_Des"
    Flag =0
    Expression ="Events.Age"
    Flag =0
    Expression ="Events.Sex"
    Flag =0
    Expression ="IIf([CompEvents].[Lane]+1=1,[House].[Lane],[CompEvents].[Lane])"
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
        dbText "Name" ="Events.Sex"
        dbInteger "ColumnWidth" ="600"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.H_NAme"
        dbInteger "ColumnWidth" ="525"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Heat"
        dbInteger "ColumnWidth" ="855"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullName"
    End
    Begin
        dbText "Name" ="A_Lane"
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
        Name ="EventType"
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
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="House"
        Name =""
    End
End
