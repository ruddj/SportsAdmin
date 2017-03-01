Operation =1
Option =8
Where ="(((EventType.Include)=True) AND ((House.Include)=True) AND ((Events.Include)=Tru"
    "e))"
Begin InputTables
    Name ="House"
    Name ="Competitors"
    Name ="CompEvents"
    Name ="Heats"
    Name ="Events"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="House.H_Code"
    Alias ="SumOfPoints"
    Expression ="Sum(CompEvents.Points)"
    Expression ="House.H_NAme"
    Expression ="House.H_ID"
    Expression ="House.CompPool"
End
Begin Joins
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =2
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
    Flag =2
End
Begin OrderBy
    Expression ="Sum(CompEvents.Points)"
    Flag =1
    Expression ="House.H_NAme"
    Flag =0
End
Begin Groups
    Expression ="House.H_Code"
    GroupLevel =0
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="House.H_ID"
    GroupLevel =0
    Expression ="House.CompPool"
    GroupLevel =0
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
        dbText "Name" ="SumOfPoints"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =66
    Top =76
    Right =962
    Bottom =513
    Left =-1
    Top =-1
    Right =878
    Bottom =243
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =188
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =444
        Top =12
        Right =540
        Bottom =119
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =585
        Top =17
        Right =681
        Bottom =124
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =710
        Top =29
        Right =806
        Bottom =136
        Top =0
        Name ="EventType"
        Name =""
    End
End
