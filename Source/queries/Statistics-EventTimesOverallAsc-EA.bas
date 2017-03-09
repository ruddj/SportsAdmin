Operation =1
Option =0
Begin InputTables
    Name ="Statistics-EventTimesOverallAsc"
End
Begin OutputColumns
    Alias ="ET_Des_Age"
    Expression ="[Statistics-EventTimesOverallAsc].ET_Des & \" - \" & [Statistics-EventTimesOvera"
        "llAsc].Age & \" \" & [Statistics-EventTimesOverallAsc].[Sex Sub]"
    Expression ="[Statistics-EventTimesOverallAsc].ET_Des"
    Expression ="[Statistics-EventTimesOverallAsc].H_NAme"
    Expression ="[Statistics-EventTimesOverallAsc].Fullname"
    Expression ="[Statistics-EventTimesOverallAsc].PIN"
    Expression ="[Statistics-EventTimesOverallAsc].[Sex Sub]"
    Expression ="[Statistics-EventTimesOverallAsc].Age"
    Expression ="[Statistics-EventTimesOverallAsc].Points"
    Expression ="[Statistics-EventTimesOverallAsc].NumericResult"
    Expression ="[Statistics-EventTimesOverallAsc].fResult"
    Expression ="[Statistics-EventTimesOverallAsc].PlaceS"
    Expression ="[Statistics-EventTimesOverallAsc].F_Lev_Sub"
    Expression ="[Statistics-EventTimesOverallAsc].F_Lev"
    Expression ="[Statistics-EventTimesOverallAsc].PlaceN"
    Expression ="[Statistics-EventTimesOverallAsc].H_Code"
    Expression ="[Statistics-EventTimesOverallAsc].Order"
    Expression ="[Statistics-EventTimesOverallAsc].OrderedResult"
End
Begin OrderBy
    Expression ="[Statistics-EventTimesOverallAsc].[Sex Sub]"
    Flag =0
    Expression ="[Statistics-EventTimesOverallAsc].ET_Des & \" - \" & [Statistics-EventTimesOvera"
        "llAsc].Age"
    Flag =0
    Expression ="[Statistics-EventTimesOverallAsc].F_Lev"
    Flag =0
    Expression ="[Statistics-EventTimesOverallAsc].PlaceN"
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
dbText "Description" ="Apply sort to Statistics-EventTimesOverallAsc with Event Age"
Begin
    Begin
        dbText "Name" ="Statistics-EventTimesOverallAsc.NumericResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Statistics-EventTimesOverallAsc.Fullname"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Statistics-EventTimesOverallAsc.fResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Statistics-EventTimesOverallAsc.PlaceS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Statistics-EventTimesOverallAsc.PlaceN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Statistics-EventTimesOverallAsc.OrderedResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ET_Des_Age"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].PlaceS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].fResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].Points"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].NumericResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].F_Lev_Sub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].PlaceN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].OrderedResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].ET_Des"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].H_NAme"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].Fullname"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].PIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].[Sex Sub]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].Age"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].F_Lev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Statistics-EventTimesOverallAsc].H_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1120
    Bottom =854
    Left =-1
    Top =-1
    Right =1102
    Bottom =312
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =426
        Top =0
        Name ="Statistics-EventTimesOverallAsc"
        Name =""
    End
End
