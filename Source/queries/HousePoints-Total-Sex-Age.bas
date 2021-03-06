﻿Operation =1
Option =8
Having ="(((House.Include)=True) AND ((House.Flag)=True) AND ((EventType.Flag)=True) AND "
    "((EventType.Include)=True))"
Begin InputTables
    Name ="CompEvents"
    Name ="House"
    Name ="Competitors"
    Name ="Events"
    Name ="Sex Sub"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="Events.Age"
    Expression ="[Sex Sub].[Sex Sub]"
    Alias ="SumOfPoints"
    Expression ="Sum(CompEvents.Points)"
    Expression ="House.H_NAme"
    Expression ="House.H_Code"
    Expression ="House.H_ID"
    Expression ="House.Include"
    Expression ="House.Flag"
    Expression ="EventType.Flag"
    Expression ="EventType.Include"
End
Begin Joins
    LeftTable ="Sex Sub"
    RightTable ="Events"
    Expression ="[Sex Sub].Sex = Events.Sex"
    Flag =3
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
    Expression ="Events.Age"
    GroupLevel =0
    Expression ="[Sex Sub].[Sex Sub]"
    GroupLevel =0
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="House.H_Code"
    GroupLevel =0
    Expression ="House.H_ID"
    GroupLevel =0
    Expression ="House.Include"
    GroupLevel =0
    Expression ="House.Flag"
    GroupLevel =0
    Expression ="EventType.Flag"
    GroupLevel =0
    Expression ="EventType.Include"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =113
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =562
        Top =41
        Right =658
        Bottom =118
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =708
        Top =6
        Right =804
        Bottom =113
        Top =0
        Name ="EventType"
        Name =""
    End
End
