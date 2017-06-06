Operation =1
Option =8
Where ="(((EventType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.Include)=Tr"
    "ue) AND ((House.Include)=True) AND ((House.Flag)=True))"
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
    Expression ="EventType.Include"
    Expression ="EventType.Flag"
    Expression ="Events.Include"
    Expression ="House.Include"
    Expression ="House.H_NAme"
    Alias ="Fullname"
    Expression ="[Surname] & \", \" & [Gname]"
    Expression ="Competitors.PIN"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="Events.Age"
    Expression ="EventType.ET_Des"
    Expression ="CompEvents.Points"
    Alias ="fResult"
    Expression ="[Result] & ' ' & [Units]"
    Expression ="CompEvents.Place"
    Expression ="[Final Level Sub].F_Lev_Sub"
    Expression ="CompEvents.F_Lev"
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
    Expression ="House.H_NAme"
    Flag =0
    Expression ="[Surname] & \", \" & [Gname]"
    Flag =0
    Expression ="CompEvents.Points"
    Flag =1
    Expression ="CompEvents.nResult"
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
        dbText "Name" ="EventType.Include"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.H_NAme"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.Include"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.PIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Sex Sub].[Sex Sub]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.Flag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events.Include"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fullname"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events.Age"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.ET_Des"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Points"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="fResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Place"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Final Level Sub].F_Lev_Sub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.F_Lev"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =81
    Top =85
    Right =1299
    Bottom =860
    Left =-1
    Top =-1
    Right =1200
    Bottom =324
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
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =325
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="Final Level Sub"
        Name =""
    End
End
