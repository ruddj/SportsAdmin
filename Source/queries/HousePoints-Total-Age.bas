﻿Operation =1
Option =8
Having ="((Not (Events.Age) Is Null) AND ((House.Include)=True) AND ((House.Flag)=True) A"
    "ND ((EventType.Flag)=True) AND ((EventType.Include)=True))"
Begin InputTables
    Name ="CompEvents"
    Name ="House"
    Name ="Competitors"
    Name ="Events"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="Events.Age"
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
    LeftTable ="CompEvents"
    RightTable ="Events"
    Expression ="CompEvents.E_Code = Events.E_Code"
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
    Expression ="Events.Age"
    Flag =0
    Expression ="Sum(CompEvents.Points)"
    Flag =1
    Expression ="House.H_NAme"
    Flag =0
End
Begin Groups
    Expression ="Events.Age"
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
    Bottom =454
    Left =-1
    Top =-1
    Right =878
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =270
        Top =12
        Right =366
        Bottom =119
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =692
        Top =2
        Right =788
        Bottom =109
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =502
        Top =7
        Right =598
        Bottom =114
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =146
        Top =9
        Right =242
        Bottom =116
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =5
        Top =14
        Right =101
        Bottom =121
        Top =0
        Name ="EventType"
        Name =""
    End
End
