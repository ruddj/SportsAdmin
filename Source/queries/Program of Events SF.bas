﻿Operation =1
Option =0
Where ="(((House.Include)=Yes) AND ((EventType.Flag)=Yes) AND ((EventType.Include)=Yes) "
    "AND ((Events.Include)=Yes))"
Begin InputTables
    Name ="EventType"
    Name ="Events"
    Name ="Heats"
    Name ="CompEvents"
    Name ="Competitors"
    Name ="Sex Sub"
    Name ="Final Level Sub"
    Name ="House"
End
Begin OutputColumns
    Expression ="House.Include"
    Expression ="EventType.Flag"
    Expression ="EventType.Include"
    Expression ="Events.Include"
    Expression ="Heats.E_Code"
    Expression ="Heats.E_Number"
    Expression ="EventType.ET_Des"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="Events.Record"
    Expression ="Events.RecName"
    Expression ="Events.RecHouse"
    Expression ="Heats.Heat"
    Expression ="[Final Level Sub].F_Lev_Sub"
    Expression ="Heats.F_Lev"
    Expression ="Heats.Status"
    Expression ="Competitors.Gname"
    Expression ="Competitors.Surname"
    Expression ="CompEvents.Lane"
    Expression ="House.H_Code"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="Sex Sub"
    Expression ="Events.Sex = [Sex Sub].Sex"
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
    Flag =1
    LeftTable ="Final Level Sub"
    RightTable ="Heats"
    Expression ="[Final Level Sub].F_Lev = Heats.F_Lev"
    Flag =1
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
    Flag =1
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
End
Begin
    State =0
    Left =69
    Top =75
    Right =1000
    Bottom =528
    Left =-1
    Top =-1
    Right =913
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =173
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =188
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =302
        Top =18
        Right =398
        Bottom =215
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =458
        Top =32
        Right =554
        Bottom =244
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =604
        Top =6
        Right =700
        Bottom =113
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =28
        Top =203
        Right =124
        Bottom =280
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =143
        Top =218
        Right =239
        Bottom =295
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =738
        Top =6
        Right =834
        Bottom =113
        Top =0
        Name ="House"
        Name =""
    End
End
