﻿Operation =1
Option =2
Having ="(((UCase([H_Code]))=FinalHouse()) AND ((House.Flag)=Yes))"
Begin InputTables
    Name ="Heats"
    Name ="EventType"
    Name ="Events"
    Name ="House"
    Name ="Temporary Table"
End
Begin OutputColumns
    Alias ="Reference#"
    Expression ="First(Heats.HE_Code)"
    Alias ="House"
    Expression ="UCase([H_Code])"
    Alias ="Gender"
    Expression ="SetSexFormat([Sex])"
    Expression ="Events.Age"
    Alias ="Event"
    Expression ="EventType.ET_Des"
    Alias ="Heats"
    Expression ="SetHeatFormat([Heat])"
    Alias ="Competitor"
    Expression ="[Temporary Table].Field1"
End
Begin Joins
    LeftTable ="EventType"
    RightTable ="Temporary Table"
    Expression ="EventType.ET_Code = [Temporary Table].ET_Code"
    Flag =1
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
    Expression ="UCase([H_Code])"
    Flag =0
    Expression ="EventType.ET_Des"
    Flag =0
    Expression ="SetSexFormat([Sex])"
    Flag =0
    Expression ="Events.Age"
    Flag =0
    Expression ="SetHeatFormat([Heat])"
    Flag =0
End
Begin Groups
    Expression ="UCase([H_Code])"
    GroupLevel =0
    Expression ="SetSexFormat([Sex])"
    GroupLevel =0
    Expression ="Events.Age"
    GroupLevel =0
    Expression ="EventType.ET_Des"
    GroupLevel =0
    Expression ="SetHeatFormat([Heat])"
    GroupLevel =0
    Expression ="[Temporary Table].Field1"
    GroupLevel =0
    Expression ="EventType.ET_Des"
    GroupLevel =0
    Expression ="Heats.E_Code"
    GroupLevel =0
    Expression ="House.Flag"
    GroupLevel =0
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
        dbText "Name" ="Events.Age"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gender"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reference#"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Heats"
    End
    Begin
        dbText "Name" ="Competitor"
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
    ColumnsShown =543
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
        Name ="EventType"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Temporary Table"
        Name =""
    End
End
