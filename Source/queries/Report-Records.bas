﻿Operation =1
Option =8
Where ="(((EventType.Include)=Yes) AND ((EventType.Flag)=Yes And (EventType.Flag)=Yes))"
Begin InputTables
    Name ="EventType"
    Name ="Events"
    Name ="Sex Sub"
    Name ="House"
    Name ="Records"
    Name ="Units"
End
Begin OutputColumns
    Alias ="BestResult"
    Expression ="IIf([Order]=\"ASC\",[nresult],1/[nResult])"
    Expression ="EventType.ET_Code"
    Expression ="Events.E_Code"
    Expression ="EventType.ET_Des"
    Expression ="EventType.Units"
    Expression ="EventType.Include"
    Expression ="EventType.Flag"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="Records.Result"
    Alias ="FullName"
    Expression ="[Gname] & \" \" & UCase([Surname])"
    Alias ="CompetitorHouse"
    Expression ="IIf(IsNull([H_NAme]),[Records].[H_Code],[H_Name])"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="EventType.Flag"
    Expression ="Records.Date"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="Sex Sub"
    Expression ="Events.Sex = [Sex Sub].Sex"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Units"
    Expression ="EventType.Units = Units.DisplayUnit"
    Flag =2
    LeftTable ="Events"
    RightTable ="Records"
    Expression ="Events.E_Code = Records.E_Code"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =2
    LeftTable ="House"
    RightTable ="Records"
    Expression ="House.H_Code = Records.H_Code"
    Flag =3
End
Begin OrderBy
    Expression ="Events.E_Code"
    Flag =0
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
        dbText "Name" ="BestResult"
    End
    Begin
        dbText "Name" ="FullName"
    End
    Begin
        dbText "Name" ="CompetitorHouse"
    End
End
Begin
    State =0
    Left =53
    Top =65
    Right =896
    Bottom =427
    Left =-1
    Top =-1
    Right =825
    Bottom =184
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =144
        Top =9
        Right =240
        Bottom =146
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =277
        Top =6
        Right =373
        Bottom =158
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =381
        Top =73
        Right =477
        Bottom =150
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =703
        Top =8
        Right =812
        Bottom =160
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =485
        Top =1
        Right =581
        Bottom =168
        Top =0
        Name ="Records"
        Name =""
    End
    Begin
        Left =17
        Top =60
        Right =113
        Bottom =167
        Top =0
        Name ="Units"
        Name =""
    End
End