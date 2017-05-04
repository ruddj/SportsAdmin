Operation =1
Option =0
Where ="(((Competitors.Gname)<>\"Team\") AND ((CompEvents.Place)<=DLookUp(\"[Mtop]\",\"M"
    "iscellaneous\")) AND ((CompEvents.F_Lev)=0) AND ((EventType.Include)=True) AND ("
    "(EventType.Flag)=True) AND ((Events.Include)=True) AND ((EventType.Mevent)<>\"\""
    "))"
Begin InputTables
    Name ="EventType"
    Name ="Competitors"
    Name ="Events"
    Name ="CompEvents"
End
Begin OutputColumns
    Alias ="EntryRecord"
    Expression ="\"D;\" & Competitors.Surname & \";\" & Competitors.Gname & \";;\" & Competitors."
        "Sex & \";\" & Format(Competitors.DOB,\"mm/dd/yy\") & \";\" & DLookUp(\"[Mcode]\""
        ",\"Miscellaneous\") & \";\" & DLookUp(\"[Mteam]\",\"Miscellaneous\") & \";;;\" &"
        " EventType.Mevent & \";\" & CompEvents.nResult & \";M;\""
End
Begin Joins
    LeftTable ="Events"
    RightTable ="CompEvents"
    Expression ="Events.E_Code = CompEvents.E_Code"
    Flag =3
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =2
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =3
End
Begin OrderBy
    Expression ="Competitors.Age"
    Flag =1
    Expression ="Competitors.Surname"
    Flag =0
    Expression ="Competitors.Gname"
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
        dbText "Name" ="EntryRecord"
        dbInteger "ColumnWidth" ="7035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1302
    Bottom =871
    Left =-1
    Top =-1
    Right =1284
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =292
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =306
        Top =0
        Name ="Competitors"
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
        Bottom =378
        Top =0
        Name ="CompEvents"
        Name =""
    End
End
