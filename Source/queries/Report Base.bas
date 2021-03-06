﻿Operation =1
Option =2
Where ="(((House.Include)=True) AND ((House.Flag)=True) AND ((EventType.Flag)=True))"
Begin InputTables
    Name ="House"
    Name ="Events"
    Name ="EventType"
    Name ="Sex Sub"
End
Begin OutputColumns
    Expression ="House.H_NAme"
    Expression ="House.H_Code"
    Expression ="House.H_ID"
    Expression ="House.Include"
    Expression ="House.Flag"
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
    Expression ="EventType.Flag"
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
    Expression ="House.H_NAme"
    Flag =0
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
    Left =106
    Top =62
    Right =1002
    Bottom =367
    Left =-1
    Top =-1
    Right =878
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =272
        Top =8
        Right =368
        Bottom =115
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =158
        Top =6
        Right =254
        Bottom =113
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =83
        Top =0
        Name ="Sex Sub"
        Name =""
    End
End
