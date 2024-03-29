﻿Operation =1
Option =0
Where ="((([ReportBase-Events].Flag)=Yes) And (([ReportBase-Events].Include)=Yes) And (("
    "[ReportBase-Events].Age) Like Forms!Reports_Event!Age_EB) And (([ReportBase-Even"
    "ts].Sex) Like Forms!Reports_Event!Sex_DD) And ((Heats.F_Lev) Like Forms!Reports_"
    "Event!Flev_DD) And ((Heats.Heat) Like Forms!Reports_Event!Heat_EB))"
Begin InputTables
    Name ="House"
    Name ="Final Level Sub"
    Name ="Heats"
    Name ="ReportBase-Events"
    Name ="Competitors"
    Name ="CompEvents"
End
Begin OutputColumns
    Expression ="[ReportBase-Events].Flag"
    Expression ="[ReportBase-Events].Include"
    Expression ="[ReportBase-Events].ET_Des"
    Expression ="[ReportBase-Events].[Sex Sub]"
    Expression ="[ReportBase-Events].Age"
    Expression ="[ReportBase-Events].Sex"
    Expression ="[ReportBase-Events].R_Code"
    Expression ="[ReportBase-Events].Record"
    Expression ="[ReportBase-Events].Units"
    Alias ="FullName"
    Expression ="IIf(IsNull([Surname]),\"\",UCase(Trim([Surname])) & \", \" & [Gname])"
    Expression ="CompEvents.Result"
    Expression ="CompEvents.nResult"
    Expression ="Heats.E_Number"
    Expression ="Heats.F_Lev"
    Expression ="Heats.Heat"
    Expression ="CompEvents.Points"
    Alias ="FLevSub"
    Expression ="IIf(IsNull([F_Lev_Sub]),[Heats].[F_Lev],[F_Lev_Sub])"
    Alias ="F_Place"
    Expression ="IIf([Place]=0,'',[Place])"
    Alias ="RecHolder"
    Expression ="DisplayRecHolder([RecName],[RecHouse])"
    Expression ="Heats.HE_Code"
    Alias ="cResult"
    Expression ="DisplayResult([Result]) & ' ' & [Units]"
    Alias ="cPoints"
    Expression ="DisplayPoints([Points])"
    Expression ="House.H_Code"
    Expression ="House.H_NAme"
    Expression ="Heats.Status"
    Expression ="Heats.E_Time"
End
Begin Joins
    LeftTable ="Heats"
    RightTable ="ReportBase-Events"
    Expression ="Heats.E_Code = [ReportBase-Events].E_Code"
    Flag =3
    LeftTable ="Final Level Sub"
    RightTable ="Heats"
    Expression ="[Final Level Sub].F_Lev = Heats.F_Lev"
    Flag =3
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
    Flag =2
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =2
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =2
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =3
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
        dbText "Name" ="FullName"
    End
    Begin
        dbText "Name" ="FLevSub"
    End
    Begin
        dbText "Name" ="F_Place"
    End
    Begin
        dbText "Name" ="RecHolder"
    End
    Begin
        dbText "Name" ="cResult"
    End
    Begin
        dbText "Name" ="cPoints"
    End
End
Begin
    State =0
    Left =78
    Top =96
    Right =798
    Bottom =472
    Left =-1
    Top =-1
    Right =702
    Bottom =194
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =227
        Top =6
        Right =323
        Bottom =113
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =24
        Top =16
        Right =120
        Bottom =123
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =374
        Top =13
        Right =470
        Bottom =120
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =488
        Top =107
        Right =584
        Bottom =184
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =593
        Top =9
        Right =689
        Bottom =191
        Top =0
        Name ="ReportBase-Events"
        Name =""
    End
    Begin
        Left =144
        Top =119
        Right =240
        Bottom =226
        Top =0
        Name ="House"
        Name =""
    End
End
