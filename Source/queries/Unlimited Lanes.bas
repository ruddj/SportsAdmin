﻿Operation =1
Option =0
Where ="(((Competitors.Sex) Like DLookUp(\"[Rsex]\",\"Miscellaneous\")) AND ((Competitor"
    "s.Age) Like DLookUp(\"[Rage]\",\"Miscellaneous\")) AND ((Heats.Heat) Like DLookU"
    "p(\"[Rheat]\",\"Miscellaneous\")) AND (([Final Level Sub].F_Lev) Like DLookUp(\""
    "[Rfinal]\",\"Miscellaneous\")))"
Begin InputTables
    Name ="House"
    Name ="EventType"
    Name ="Events"
    Name ="Final Level Sub"
    Name ="Heats"
    Name ="Competitors"
    Name ="CompEvents"
    Name ="ReportTypes"
End
Begin OutputColumns
    Alias ="FullName"
    Expression ="DetermineFullName([Surname],[Gname])"
    Expression ="Competitors.H_Code"
    Alias ="F_Place"
    Expression ="IIf([Place]=0,'',[Place])"
    Alias ="cResult"
    Expression ="DisplayResult([Result]) & ' ' & [Units]"
    Expression ="CompEvents.Memo"
    Alias ="cPoints"
    Expression ="DisplayPoints([Points])"
    Alias ="RecHolder"
    Expression ="DisplayRecHolder([RecName],[RecHouse])"
    Expression ="Competitors.Sex"
    Expression ="Competitors.Age"
    Expression ="Heats.E_Number"
    Alias ="FLevSub"
    Expression ="IIf(IsNull([F_Lev_Sub]),[Heats].[F_Lev],[F_Lev_Sub])"
    Expression ="EventType.Units"
    Expression ="EventType.ET_Des"
    Expression ="Events.Record"
    Expression ="Heats.Heat"
    Expression ="[Final Level Sub].F_Lev"
    Expression ="ReportTypes.R_Code"
    Expression ="Heats.HE_Code"
End
Begin Joins
    LeftTable ="ReportTypes"
    RightTable ="EventType"
    Expression ="ReportTypes.R_Code = EventType.R_Code"
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
    Begin
        dbText "Name" ="F_Place"
    End
    Begin
        dbText "Name" ="cResult"
    End
    Begin
        dbText "Name" ="cPoints"
    End
    Begin
        dbText "Name" ="RecHolder"
    End
    Begin
        dbText "Name" ="FLevSub"
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
        Name ="House"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =624
        Top =156
        Right =768
        Bottom =300
        Top =0
        Name ="ReportTypes"
        Name =""
    End
End
