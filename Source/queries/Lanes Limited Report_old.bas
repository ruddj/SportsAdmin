﻿Operation =1
Option =0
Where ="(((House.Include)=Yes) AND (([Lanes Limited].Age) Like DLookUp(\"[Rage]\",\"Misc"
    "ellaneous\")) AND (([Lanes Limited].Sex) Like DLookUp(\"[Rsex]\",\"Miscellaneous"
    "\")) AND (([Lanes Limited].F_Lev) Like DLookUp(\"[Rfinal]\",\"Miscellaneous\")) "
    "AND (([Lanes Limited].Heat) Like DLookUp(\"[Rheat]\",\"Miscellaneous\")))"
Begin InputTables
    Name ="Lanes Limited"
    Name ="CompEvents"
    Name ="Competitors"
    Name ="Final Level Sub"
    Name ="House"
End
Begin OutputColumns
    Expression ="House.Include"
    Expression ="[Lanes Limited].ET_Des"
    Expression ="[Lanes Limited].Age"
    Expression ="[Lanes Limited].Sex"
    Expression ="[Lanes Limited].F_Lev"
    Alias ="FLevSub"
    Expression ="IIf(IsNull([F_Lev_Sub]),[Heats].[F_Lev],[F_Lev_Sub])"
    Expression ="[Lanes Limited].Heat"
    Expression ="[Lanes Limited].Lanes"
    Alias ="FullName"
    Expression ="DetermineFullName([Surname],[Gname])"
    Expression ="Competitors.H_Code"
    Expression ="[Lanes Limited].E_Number"
    Alias ="F_Place"
    Expression ="IIf([Place]=0,'',[Place])"
    Alias ="cResult"
    Expression ="IIf(IsNull([Result]),DisplayResult([Result]),DisplayResult([Result]) & \" \" & ["
        "Units])"
    Expression ="[Lanes Limited].R_Code"
    Expression ="CompEvents.Memo"
    Expression ="[Lanes Limited].HE_Code"
    Alias ="cPoints"
    Expression ="DisplayPoints([Points])"
    Expression ="[Lanes Limited].nRecord"
    Expression ="[Lanes Limited].Record"
    Alias ="RecHolder"
    Expression ="DisplayRecHolder([RecName],[RecHouse])"
    Expression ="[Lanes Limited].Units"
    Alias ="LaneSub"
    Expression ="IIf(IsNull([Lane_Sub]),[Lanes],[Lane_Sub])"
    Expression ="[Lanes Limited].E_Time"
End
Begin Joins
    LeftTable ="Lanes Limited"
    RightTable ="CompEvents"
    Expression ="[Lanes Limited].F_Lev = CompEvents.F_Lev"
    Flag =2
    LeftTable ="Lanes Limited"
    RightTable ="CompEvents"
    Expression ="[Lanes Limited].Heat = CompEvents.Heat"
    Flag =2
    LeftTable ="Lanes Limited"
    RightTable ="CompEvents"
    Expression ="[Lanes Limited].E_Code = CompEvents.E_Code"
    Flag =2
    LeftTable ="Lanes Limited"
    RightTable ="CompEvents"
    Expression ="[Lanes Limited].Lanes = CompEvents.Lane"
    Flag =2
    LeftTable ="Lanes Limited"
    RightTable ="Final Level Sub"
    Expression ="[Lanes Limited].F_Lev = [Final Level Sub].F_Lev"
    Flag =2
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =3
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =3
End
Begin OrderBy
    Expression ="[Lanes Limited].ET_Des"
    Flag =0
    Expression ="[Lanes Limited].Age"
    Flag =0
    Expression ="[Lanes Limited].Sex"
    Flag =0
    Expression ="[Lanes Limited].F_Lev"
    Flag =0
    Expression ="[Lanes Limited].Heat"
    Flag =0
    Expression ="[Lanes Limited].Lanes"
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
        dbText "Name" ="[Lanes Limited].ET_Des"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Lanes Limited].Age"
        dbInteger "ColumnWidth" ="720"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Lanes Limited].Sex"
        dbInteger "ColumnWidth" ="435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecHolder"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FLevSub"
    End
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
        dbText "Name" ="LaneSub"
    End
End
Begin
    State =0
    Left =52
    Top =51
    Right =840
    Bottom =451
    Left =-1
    Top =-1
    Right =770
    Bottom =231
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =153
        Top =35
        Right =249
        Bottom =202
        Top =0
        Name ="Lanes Limited"
        Name =""
    End
    Begin
        Left =344
        Top =5
        Right =440
        Bottom =202
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =482
        Top =16
        Right =578
        Bottom =228
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =0
        Top =24
        Right =96
        Bottom =101
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =616
        Top =6
        Right =712
        Bottom =113
        Top =0
        Name ="House"
        Name =""
    End
End
