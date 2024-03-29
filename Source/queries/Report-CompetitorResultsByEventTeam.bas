﻿Operation =1
Option =8
Where ="(((CompEvents.Result) Is Not Null) AND ((EventType.Include)=Yes) AND ((EventType"
    ".Flag)=Yes) AND ((Events.Include)=Yes) AND ((House.Include)=Yes) AND ((House.Fla"
    "g)=Yes))"
Begin InputTables
    Name ="House"
    Name ="EventType"
    Name ="Competitors"
    Name ="Sex Sub"
    Name ="Events"
    Name ="CompEvents"
    Name ="Final Level Sub"
End
Begin OutputColumns
    Expression ="Events.Age"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="House.H_NAme"
    Alias ="Fullname"
    Expression ="[Surname] & \", \" & [Gname]"
    Expression ="Competitors.PIN"
    Expression ="EventType.ET_Des"
    Expression ="CompEvents.Points"
    Alias ="fResult"
    Expression ="[Result] & ' ' & [Units]"
    Expression ="CompEvents.Place"
    Expression ="CompEvents.nResult"
    Expression ="CompEvents.Result"
    Expression ="[Final Level Sub].F_Lev_Sub"
    Expression ="[Final Level Sub].F_Lev"
End
Begin Joins
    LeftTable ="Competitors"
    RightTable ="Sex Sub"
    Expression ="Competitors.Sex = [Sex Sub].Sex"
    Flag =2
    LeftTable ="Events"
    RightTable ="CompEvents"
    Expression ="Events.E_Code = CompEvents.E_Code"
    Flag =3
    LeftTable ="CompEvents"
    RightTable ="Final Level Sub"
    Expression ="CompEvents.F_Lev = [Final Level Sub].F_Lev"
    Flag =2
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
    Flag =3
End
Begin OrderBy
    Expression ="Events.Age"
    Flag =0
    Expression ="[Sex Sub].[Sex Sub]"
    Flag =0
    Expression ="House.H_NAme"
    Flag =0
    Expression ="EventType.ET_Des"
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
        dbText "Name" ="fResult"
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
    Bottom =258
    Left =288
    Top =0
    ColumnsShown =539
    Begin
        Left =610
        Top =95
        Right =706
        Bottom =202
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =-219
        Top =7
        Right =-123
        Bottom =114
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =307
        Top =13
        Right =403
        Bottom =165
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =457
        Top =9
        Right =553
        Bottom =86
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =-66
        Top =10
        Right =30
        Bottom =117
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =180
        Top =8
        Right =276
        Bottom =190
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =60
        Top =87
        Right =156
        Bottom =164
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
End
