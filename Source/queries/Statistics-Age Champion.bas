﻿Operation =1
Option =8
Having ="(((Competitors.Age) Is Not Null) AND ((EventType.Flag)=Yes) AND ((House.Flag)=Ye"
    "s) AND ((UCase([Gname]))<>\"TEAM\"))"
Begin InputTables
    Name ="CompetitorEventAge"
    Name ="House"
    Name ="Sex Sub"
    Name ="Events"
    Name ="Heats"
    Name ="Competitors"
    Name ="CompEvents"
    Name ="EventType"
End
Begin OutputColumns
    Alias ="Fullname"
    Expression ="Trim(UCase([Surname])) & \", \" & [Gname] & \" (\" & [Competitors].[Age] & \")\""
    Alias ="AgeSex"
    Expression ="[CompetitorEventAge].[Eage] & ' ' & [Sex Sub].[Sex Sub]"
    Expression ="House.H_NAme"
    Alias ="SumOfPoints"
    Expression ="Sum(CompEvents.Points)"
    Expression ="Competitors.Age"
    Expression ="EventType.Flag"
    Expression ="House.Flag"
End
Begin Joins
    LeftTable ="Sex Sub"
    RightTable ="Events"
    Expression ="[Sex Sub].Sex = Events.Sex"
    Flag =3
    LeftTable ="CompetitorEventAge"
    RightTable ="Competitors"
    Expression ="CompetitorEventAge.Cage = Competitors.Age"
    Flag =1
    LeftTable ="Events"
    RightTable ="CompetitorEventAge"
    Expression ="Events.Age = CompetitorEventAge.Eage"
    Flag =1
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =1
    LeftTable ="Events"
    RightTable ="Heats"
    Expression ="Events.E_Code = Heats.E_Code"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
    Flag =1
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =3
End
Begin OrderBy
    Expression ="[CompetitorEventAge].[Eage] & ' ' & [Sex Sub].[Sex Sub]"
    Flag =0
    Expression ="Sum(CompEvents.Points)"
    Flag =1
End
Begin Groups
    Expression ="Trim(UCase([Surname])) & \", \" & [Gname] & \" (\" & [Competitors].[Age] & \")\""
    GroupLevel =0
    Expression ="[CompetitorEventAge].[Eage] & ' ' & [Sex Sub].[Sex Sub]"
    GroupLevel =0
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="Competitors.Age"
    GroupLevel =0
    Expression ="EventType.Flag"
    GroupLevel =0
    Expression ="House.Flag"
    GroupLevel =0
    Expression ="UCase([Gname])"
    GroupLevel =0
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
        dbText "Name" ="Fullname"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgeSex"
    End
    Begin
        dbText "Name" ="SumOfPoints"
    End
End
Begin
    State =0
    Left =142
    Top =140
    Right =1405
    Bottom =663
    Left =-1
    Top =-1
    Right =1245
    Bottom =345
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =441
        Top =235
        Right =537
        Bottom =342
        Top =0
        Name ="CompetitorEventAge"
        Name =""
    End
    Begin
        Left =31
        Top =24
        Right =127
        Bottom =131
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =846
        Top =184
        Right =942
        Bottom =261
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =691
        Top =47
        Right =825
        Bottom =229
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =455
        Top =49
        Right =551
        Bottom =156
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =170
        Top =35
        Right =266
        Bottom =262
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =322
        Top =18
        Right =418
        Bottom =230
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =893
        Top =38
        Right =989
        Bottom =190
        Top =0
        Name ="EventType"
        Name =""
    End
End
