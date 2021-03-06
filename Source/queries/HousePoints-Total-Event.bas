﻿Operation =1
Option =8
Having ="(((House.Include)=True) AND ((EventType.Flag)=True))"
Begin InputTables
    Name ="CompEvents"
    Name ="House"
    Name ="Competitors"
    Name ="Events"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="EventType.ET_Des"
    Alias ="SumOfPoints"
    Expression ="Sum(CompEvents.Points)"
    Expression ="House.H_NAme"
    Expression ="House.H_Code"
    Expression ="House.H_ID"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="CompEvents"
    Expression ="Events.E_Code = CompEvents.E_Code"
    Flag =3
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =3
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =2
End
Begin OrderBy
    Expression ="Sum(CompEvents.Points)"
    Flag =1
    Expression ="House.H_NAme"
    Flag =0
End
Begin Groups
    Expression ="EventType.ET_Des"
    GroupLevel =0
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="House.H_Code"
    GroupLevel =0
    Expression ="House.H_ID"
    GroupLevel =0
    Expression ="House.Include"
    GroupLevel =0
    Expression ="EventType.Flag"
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
        dbText "Name" ="SumOfPoints"
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
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="EventType"
        Name =""
    End
End
