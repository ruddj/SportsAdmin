﻿Operation =1
Option =8
Where ="(((EventType.Flag)=Yes))"
Begin InputTables
    Name ="EventType"
    Name ="Events"
    Name ="Sex Sub"
    Name ="House"
    Name ="Temporary Record-Best in Full"
End
Begin OutputColumns
    Expression ="EventType.ET_Code"
    Expression ="EventType.ET_Des"
    Expression ="EventType.Units"
    Expression ="EventType.Include"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="Events.Record"
    Alias ="FullName"
    Expression ="[Gname] & \" \" & UCase([Surname])"
    Expression ="Events.RecHouse"
    Expression ="House.H_NAme"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="EventType.Flag"
    Expression ="[Temporary Record-Best in Full].Date"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="Sex Sub"
    Expression ="Events.Sex = [Sex Sub].Sex"
    Flag =2
    LeftTable ="Events"
    RightTable ="House"
    Expression ="Events.RecHouse = House.H_ID"
    Flag =2
    LeftTable ="Events"
    RightTable ="Temporary Record-Best in Full"
    Expression ="Events.E_Code = [Temporary Record-Best in Full].E_Code"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FullName"
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
        Name ="Sex Sub"
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
        Name ="Temporary Record-Best in Full"
        Name =""
    End
End
