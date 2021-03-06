﻿Operation =1
Option =8
Where ="(((EventType.Flag)=True) And ((EventType.Include)=True) And ((Events.Include)=Tr"
    "ue) And ((Events.Age) Like Forms!Reports_Event!Age_EB) And ((Events.Sex) Like Fo"
    "rms!Reports_Event!Sex_DD) And ((Heats.F_Lev) Like Forms!Reports_Event!Flev_DD) A"
    "nd ((Heats.Heat) Like Forms!Reports_Event!Heat_EB) And ((Heats.Status) Like Form"
    "s!Reports_Event!FutureEB Or (Heats.Status) Like Forms!Reports_Event!ActiveEB Or "
    "(Heats.Status) Like Forms!Reports_Event!CompletedEB Or (Heats.Status) Like Forms"
    "!Reports_Event!PromotedEB))"
Begin InputTables
    Name ="EventType"
    Name ="Lane Sub"
    Name ="Lane Template"
    Name ="Events"
    Name ="Heats"
End
Begin OutputColumns
    Expression ="EventType.Flag"
    Expression ="EventType.Include"
    Expression ="Events.Include"
    Expression ="Heats.HE_Code"
    Expression ="Events.ET_Code"
    Expression ="EventType.R_Code"
    Expression ="Events.E_Code"
    Expression ="EventType.ET_Des"
    Expression ="Events.Age"
    Expression ="Events.Sex"
    Expression ="Heats.F_Lev"
    Expression ="Heats.Heat"
    Expression ="[Lane Template].Lanes"
    Expression ="Heats.E_Number"
    Expression ="Heats.Status"
    Expression ="Events.nRecord"
    Expression ="Events.Record"
    Expression ="Events.RecName"
    Expression ="Events.RecHouse"
    Expression ="EventType.Units"
    Expression ="[Lane Sub].Lane_Sub"
    Expression ="Heats.E_Time"
End
Begin Joins
    LeftTable ="Lane Sub"
    RightTable ="Lane Template"
    Expression ="[Lane Sub].Lane = [Lane Template].Lanes"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Lane Template"
    Expression ="EventType.ET_Code = [Lane Template].ET_Code"
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
    Expression ="Events.ET_Code"
    Flag =0
    Expression ="Events.Age"
    Flag =0
    Expression ="Events.Sex"
    Flag =0
    Expression ="Heats.F_Lev"
    Flag =0
    Expression ="Heats.Heat"
    Flag =0
    Expression ="[Lane Template].Lanes"
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
        dbText "Name" ="EventType.Include"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Heats.HE_Code"
        dbInteger "ColumnWidth" ="1215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events.E_Code"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.ET_Des"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =10
    Right =918
    Bottom =315
    Left =-1
    Top =-1
    Right =794
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =139
        Top =7
        Right =235
        Bottom =84
        Top =0
        Name ="Lane Template"
        Name =""
    End
    Begin
        Left =439
        Top =9
        Right =535
        Bottom =116
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =593
        Top =8
        Right =689
        Bottom =115
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =303
        Top =7
        Right =399
        Bottom =114
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =10
        Top =15
        Right =106
        Bottom =92
        Top =0
        Name ="Lane Sub"
        Name =""
    End
End
