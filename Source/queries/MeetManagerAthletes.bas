Operation =1
Option =8
Where ="(((Competitors.Gname)<>\"Team\") AND (([Statistics-CompetitorEvents].F_Lev)=0) A"
    "ND (([Statistics-CompetitorEvents].Place)<=DLookUp(\"[Mtop]\",\"Miscellaneous\")"
    "))"
Begin InputTables
    Name ="Statistics-CompetitorEvents"
    Name ="Competitors"
End
Begin OutputColumns
    Alias ="InfoRecord"
    Expression ="\"I;\" & Competitors.Surname & \";\" & Competitors.Gname & \";;\" & Competitors."
        "Sex & \";\" & Format(Competitors.DOB,\"mm/dd/yy\") & \";\" & DLookUp(\"[Mcode]\""
        ",\"Miscellaneous\") & \";\" & DLookUp(\"[Mteam]\",\"Miscellaneous\") & \";\""
End
Begin Joins
    LeftTable ="Statistics-CompetitorEvents"
    RightTable ="Competitors"
    Expression ="[Statistics-CompetitorEvents].PIN = Competitors.PIN"
    Flag =1
End
Begin OrderBy
    Expression ="\"I;\" & Competitors.Surname & \";\" & Competitors.Gname & \";;\" & Competitors."
        "Sex & \";\" & Format(Competitors.DOB,\"mm/dd/yy\") & \";Code;Name;\""
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
        dbText "Name" ="InfoRecord"
        dbInteger "ColumnWidth" ="6270"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =62
    Top =54
    Right =1320
    Bottom =832
    Left =-1
    Top =-1
    Right =1240
    Bottom =416
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =391
        Top =0
        Name ="Statistics-CompetitorEvents"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =380
        Top =0
        Name ="Competitors"
        Name =""
    End
End
