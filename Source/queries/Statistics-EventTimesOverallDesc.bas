﻿Operation =1
Option =8
Where ="(((Competitors.PIN) Is Not Null) AND ((EventType.Flag)=Yes) AND ((Units.Order)=\""
    "Desc\"))"
Begin InputTables
    Name ="House"
    Name ="EventType"
    Name ="Competitors"
    Name ="Sex Sub"
    Name ="Events"
    Name ="CompEvents"
    Name ="Final Level Sub"
    Name ="Units"
End
Begin OutputColumns
    Expression ="EventType.ET_Des"
    Expression ="House.H_NAme"
    Alias ="Fullname"
    Expression ="[Surname] & \", \" & [Gname]"
    Expression ="Competitors.PIN"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="Events.Age"
    Expression ="CompEvents.Points"
    Alias ="NumericResult"
    Expression ="IIf([nResult]=0,1E+31,[nresult])"
    Alias ="fResult"
    Expression ="[Result] & ' ' & [Units]"
    Alias ="PlaceS"
    Expression ="IIf([Place]=0,'-',Str([Place]))"
    Expression ="[Final Level Sub].F_Lev_Sub"
    Expression ="[Final Level Sub].F_Lev"
    Alias ="PlaceN"
    Expression ="IIf([Place]=0,1E+31,[Place])"
    Expression ="House.H_Code"
    Expression ="EventType.Flag"
    Expression ="Units.Order"
End
Begin Joins
    LeftTable ="Competitors"
    RightTable ="Sex Sub"
    Expression ="Competitors.Sex = [Sex Sub].Sex"
    Flag =2
    LeftTable ="Events"
    RightTable ="CompEvents"
    Expression ="Events.E_Code = CompEvents.E_Code"
    Flag =2
    LeftTable ="CompEvents"
    RightTable ="Final Level Sub"
    Expression ="CompEvents.F_Lev = [Final Level Sub].F_Lev"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Units"
    Expression ="EventType.Units = Units.DisplayUnit"
    Flag =2
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =3
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =2
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =3
End
Begin OrderBy
    Expression ="[Surname] & \", \" & [Gname]"
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
        dbText "Name" ="Fullname"
    End
    Begin
        dbText "Name" ="NumericResult"
    End
    Begin
        dbText "Name" ="fResult"
    End
    Begin
        dbText "Name" ="PlaceS"
    End
    Begin
        dbText "Name" ="PlaceN"
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
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =624
        Top =156
        Right =768
        Bottom =300
        Top =0
        Name ="Units"
        Name =""
    End
End
