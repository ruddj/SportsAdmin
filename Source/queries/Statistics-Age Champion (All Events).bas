Operation =1
Option =8
Having ="(((Competitors.Age) Is Not Null) AND ((UCase([Gname]))<>\"TEAM\"))"
Begin InputTables
    Name ="CompetitorEventAge"
    Name ="House"
    Name ="Sex Sub"
    Name ="Events"
    Name ="Heats"
    Name ="Competitors"
    Name ="CompEvents"
    Name ="EventType"
End
Begin OutputColumns
    Alias ="Fullname"
    Expression ="Trim(UCase([Surname])) & \", \" & [Gname]"
    Alias ="AgeSex"
    Expression ="[CompetitorEventAge].[Eage] & ' ' & [Sex Sub].[Sex Sub]"
    Expression ="House.H_NAme"
    Alias ="SumOfPoints"
    Expression ="Sum(CompEvents.Points)"
    Expression ="Competitors.Age"
End
Begin Joins
    LeftTable ="Sex Sub"
    RightTable ="Events"
    Expression ="[Sex Sub].Sex = Events.Sex"
    Flag =3
    LeftTable ="CompetitorEventAge"
    RightTable ="Competitors"
    Expression ="CompetitorEventAge.Cage = Competitors.Age"
    Flag =3
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =3
    LeftTable ="Events"
    RightTable ="Heats"
    Expression ="Events.E_Code = Heats.E_Code"
    Flag =3
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =3
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
    Flag =3
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =3
End
Begin OrderBy
    Expression ="[CompetitorEventAge].[Eage] & ' ' & [Sex Sub].[Sex Sub]"
    Flag =0
    Expression ="Sum(CompEvents.Points)"
    Flag =1
End
Begin Groups
    Expression ="Trim(UCase([Surname])) & \", \" & [Gname]"
    GroupLevel =0
    Expression ="[CompetitorEventAge].[Eage] & ' ' & [Sex Sub].[Sex Sub]"
    GroupLevel =0
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="Competitors.Age"
    GroupLevel =0
    Expression ="UCase([Gname])"
    GroupLevel =0
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
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgeSex"
    End
    Begin
        dbText "Name" ="SumOfPoints"
    End
End
Begin
    State =0
    Left =87
    Top =62
    Right =974
    Bottom =595
    Left =-1
    Top =-1
    Right =869
    Bottom =345
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =32
        Top =146
        Right =128
        Bottom =223
        Top =0
        Name ="CompetitorEventAge"
        Name =""
    End
    Begin
        Left =31
        Top =24
        Right =127
        Bottom =131
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =729
        Top =70
        Right =825
        Bottom =147
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =587
        Top =49
        Right =683
        Bottom =156
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =455
        Top =49
        Right =551
        Bottom =156
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =170
        Top =35
        Right =266
        Bottom =187
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =322
        Top =18
        Right =418
        Bottom =125
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =457
        Top =199
        Right =553
        Bottom =306
        Top =0
        Name ="EventType"
        Name =""
    End
End
