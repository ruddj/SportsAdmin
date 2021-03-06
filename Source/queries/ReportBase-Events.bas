﻿Operation =1
Option =2
Where ="(((House.Include)=True))"
Begin InputTables
    Name ="House"
    Name ="Events"
    Name ="EventType"
    Name ="Sex Sub"
End
Begin OutputColumns
    Expression ="House.Include"
    Expression ="Events.Age"
    Expression ="Events.Sex"
    Expression ="EventType.ET_Des"
    Expression ="EventType.Flag"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="EventType.R_Code"
    Expression ="Events.E_Code"
    Expression ="EventType.ET_Code"
    Expression ="EventType.Units"
    Expression ="Events.Record"
    Expression ="Events.RecName"
    Expression ="Events.RecHouse"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="Sex Sub"
    Expression ="Events.Sex = [Sex Sub].Sex"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =1
End
Begin OrderBy
    Expression ="Events.Age"
    Flag =0
    Expression ="Events.Sex"
    Flag =0
    Expression ="EventType.ET_Des"
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
        dbText "Name" ="EventType.ET_Des"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.Flag"
        dbInteger "ColumnWidth" ="690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Sex Sub].[Sex Sub]"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.R_Code"
        dbInteger "ColumnWidth" ="750"
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
        Name ="House"
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
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Sex Sub"
        Name =""
    End
End
