Operation =1
Option =2
Where ="(((Competitors.Gname)<>\"Team\") AND (([Statistics-CompetitorEvents].F_Lev)=0) A"
    "ND (([Statistics-CompetitorEvents].Place)<=DLookUp(\"[Mtop]\",\"Miscellaneous\")"
    "))"
Begin InputTables
    Name ="Statistics-CompetitorEvents"
    Name ="Competitors"
End
Begin OutputColumns
    Expression ="Competitors.ID"
    Alias ="Surname"
    Expression ="Left([Competitors].[Surname],20)"
    Alias ="Given"
    Expression ="Left([Competitors].[Gname],20)"
    Expression ="Competitors.Sex"
    Expression ="Competitors.DOB"
End
Begin Joins
    LeftTable ="Statistics-CompetitorEvents"
    RightTable ="Competitors"
    Expression ="[Statistics-CompetitorEvents].PIN = Competitors.PIN"
    Flag =1
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
        dbText "Name" ="Competitors.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Surname"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Given"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.Sex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.DOB"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =26
    Top =96
    Right =963
    Bottom =874
    Left =-1
    Top =-1
    Right =919
    Bottom =281
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
