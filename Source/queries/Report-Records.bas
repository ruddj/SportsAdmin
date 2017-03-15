Operation =1
Option =8
Where ="(((EventType.Include)=Yes) AND ((EventType.Flag)=Yes And (EventType.Flag)=Yes))"
Begin InputTables
    Name ="House"
    Name ="EventType"
    Name ="Units"
    Name ="Events"
    Name ="Sex Sub"
    Name ="Records"
End
Begin OutputColumns
    Expression ="EventType.ET_Des"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Alias ="BestResult"
    Expression ="IIf([Order]=\"ASC\",[nresult],1/[nResult])"
    Expression ="EventType.ET_Code"
    Expression ="Events.E_Code"
    Expression ="EventType.Units"
    Expression ="EventType.Include"
    Expression ="EventType.Flag"
    Alias ="AgeSex"
    Expression ="Events.Age & \" \" & [Sex Sub].[Sex Sub]"
    Expression ="Records.Result"
    Alias ="ResultFormated"
    Expression ="Records.Result & \" \" & EventType.Units"
    Alias ="FullName"
    Expression ="[Gname] & \" \" & UCase([Surname])"
    Alias ="CompetitorHouse"
    Expression ="IIf(IsNull([H_NAme]),[Records].[H_Code],[H_Name])"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="EventType.Flag"
    Expression ="Records.Date"
End
Begin Joins
    LeftTable ="EventType"
    RightTable ="Units"
    Expression ="EventType.Units = Units.DisplayUnit"
    Flag =2
    LeftTable ="Events"
    RightTable ="Sex Sub"
    Expression ="Events.Sex = [Sex Sub].Sex"
    Flag =2
    LeftTable ="Events"
    RightTable ="Records"
    Expression ="Events.E_Code = Records.E_Code"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =2
    LeftTable ="House"
    RightTable ="Records"
    Expression ="House.H_Code = Records.H_Code"
    Flag =3
End
Begin OrderBy
    Expression ="EventType.ET_Des"
    Flag =0
    Expression ="Events.Sex"
    Flag =1
    Expression ="Events.Age"
    Flag =0
    Expression ="IIf([Order]=\"ASC\",[nresult],1/[nResult])"
    Flag =0
    Expression ="Events.E_Code"
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
        dbText "Name" ="BestResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompetitorHouse"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.ET_Des"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.Include"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.ET_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events.E_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events.Sex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Events.Age"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.Result"
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
        dbText "Name" ="Records.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgeSex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ResultFormated"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =56
    Top =92
    Right =1202
    Bottom =621
    Left =-1
    Top =-1
    Right =1128
    Bottom =133
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =703
        Top =8
        Right =812
        Bottom =160
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =144
        Top =9
        Right =240
        Bottom =146
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =17
        Top =60
        Right =113
        Bottom =167
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =277
        Top =6
        Right =373
        Bottom =158
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =381
        Top =73
        Right =477
        Bottom =150
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =485
        Top =1
        Right =581
        Bottom =168
        Top =0
        Name ="Records"
        Name =""
    End
End
