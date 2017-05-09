Operation =1
Option =2
Begin InputTables
    Name ="CompetitorEventAge"
End
Begin OutputColumns
    Expression ="CompetitorEventAge.Eage"
    Expression ="CompetitorEventAge.Mdiv"
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
        dbText "Name" ="CompetitorEventAge.Eage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompetitorEventAge.Mdiv"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =44
    Top =93
    Right =1302
    Bottom =871
    Left =-1
    Top =-1
    Right =1240
    Bottom =501
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="CompetitorEventAge"
        Name =""
    End
End
