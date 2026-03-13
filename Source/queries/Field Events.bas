Operation =1
Option =0
Where ="((([ReportBase-Events].Flag)=Yes) AND (([ReportBase-Events].Include)=Yes) AND (("
    "[ReportBase-Events].Age) Like [Forms]![Reports_Event]![Age_EB]) AND (([ReportBas"
    "e-Events].Sex) Like [Forms]![Reports_Event]![Sex_DD]) AND ((Heats.F_Lev) Like [F"
    "orms]![Reports_Event]![Flev_DD]) AND ((Heats.Heat) Like [Forms]![Reports_Event]!"
    "[Heat_EB]))"
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
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =3
    LeftTable ="Final Level Sub"
    RightTable ="Heats"
    Expression ="[Final Level Sub].F_Lev = Heats.F_Lev"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =2
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =2
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FLevSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="F_Place"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecHolder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cPoints"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[ReportBase-Events].Flag"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =96
    Right =1022
    Bottom =859
    Left =-1
    Top =-1
    Right =920
    Bottom =343
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =-3
        Top =74
        Right =93
        Bottom =306
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =559
        Top =231
        Right =695
        Bottom =326
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
    Begin
        Left =429
        Top =3
        Right =525
        Bottom =322
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =591
        Top =10
        Right =687
        Bottom =192
        Top =0
        Name ="ReportBase-Events"
        Name =""
    End
    Begin
        Left =129
        Top =9
        Right =225
        Bottom =310
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =265
        Top =10
        Right =361
        Bottom =272
        Top =0
        Name ="CompEvents"
        Name =""
    End
End
