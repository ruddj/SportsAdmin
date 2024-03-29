﻿Operation =1
Option =8
Begin InputTables
    Name ="EventType"
    Name ="Events"
    Name ="Heats"
End
Begin OutputColumns
    Expression ="Heats.HE_Code"
    Expression ="Events.ET_Code"
    Expression ="EventType.ET_Des"
    Expression ="Events.Age"
    Expression ="Events.Sex"
    Expression ="Heats.F_Lev"
    Expression ="Heats.Heat"
    Expression ="EventType.Units"
    Expression ="Events.E_Code"
    Alias ="join"
    Expression ="Str([Events].[E_Code])+Str([Heats].[Heat])"
    Expression ="Events.Record"
    Expression ="Events.nRecord"
    Expression ="Heats.E_Number"
    Expression ="Heats.Completed"
    Expression ="Heats.Status"
    Expression ="EventType.Lane_Cnt"
    Expression ="Heats.AllNames"
    Expression ="Heats.PtScale"
    Expression ="EventType.PlacesAcrossAllHeats"
    Expression ="Heats.DontOverridePlaces"
    Expression ="Heats.EffectsRecords"
End
Begin Joins
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
    Expression ="EventType.ET_Des"
    Flag =0
    Expression ="Events.Age"
    Flag =0
    Expression ="Events.Sex"
    Flag =0
    Expression ="Heats.F_Lev"
    Flag =0
    Expression ="Heats.Heat"
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
        dbText "Name" ="join"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =169
    Top =166
    Right =1056
    Bottom =471
    Left =-1
    Top =-1
    Right =869
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
        Name ="EventType"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="Heats"
        Name =""
    End
End
