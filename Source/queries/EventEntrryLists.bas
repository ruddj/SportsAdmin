﻿Operation =1
Option =8
Where ="(((Events.Sex) Like Forms!Reports_Event!Sex_DD) And ((Events.Age) Like Forms!Rep"
    "orts_Event!Age_EB) And ((Heats.Status) Like Forms!Reports_Event!FutureEB Or (Hea"
    "ts.Status) Like Forms!Reports_Event!ActiveEB Or (Heats.Status) Like Forms!Report"
    "s_Event!CompletedEB Or (Heats.Status) Like Forms!Reports_Event!PromotedEB) And ("
    "(EventType.Flag)=True) And ((EventType.Include)=True))"
Begin InputTables
    Name ="Final Level Sub"
    Name ="EventType"
    Name ="Events"
    Name ="Heats"
End
Begin OutputColumns
    Alias ="E_Num"
    Expression ="\"# \"+Trim(Str([E_Number]))"
    Expression ="EventType.ET_Des"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="Heats.Heat"
    Expression ="Heats.E_Number"
    Expression ="[Final Level Sub].F_Lev_Sub"
    Expression ="Heats.Status"
    Expression ="EventType.Flag"
    Expression ="EventType.Include"
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
    LeftTable ="Final Level Sub"
    RightTable ="Heats"
    Expression ="[Final Level Sub].F_Lev = Heats.F_Lev"
    Flag =1
End
Begin OrderBy
    Expression ="\"# \"+Trim(Str([E_Number]))"
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
        dbText "Name" ="E_Num"
    End
End
Begin
    State =0
    Left =71
    Top =70
    Right =1002
    Bottom =372
    Left =-1
    Top =-1
    Right =913
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
        Bottom =83
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =113
        Top =0
        Name ="Heats"
        Name =""
    End
End
