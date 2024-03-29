﻿Operation =1
Option =8
Begin InputTables
    Name ="CompEvents"
    Name ="Competitors"
    Name ="Heats"
End
Begin OutputColumns
    Expression ="CompEvents.PIN"
    Expression ="CompEvents.E_Code"
    Expression ="CompEvents.Heat"
    Expression ="CompEvents.Result"
    Expression ="CompEvents.nResult"
    Expression ="CompEvents.F_Lev"
    Expression ="Competitors.Gname"
    Expression ="Competitors.Surname"
    Expression ="Competitors.H_Code"
    Expression ="Competitors.H_ID"
    Expression ="Heats.EffectsRecords"
End
Begin Joins
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
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
        dbText "Name" ="CompEvents.PIN"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.E_Code"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Heat"
        dbInteger "ColumnWidth" ="690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Result"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.nResult"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.F_Lev"
        dbInteger "ColumnWidth" ="570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =84
    Top =40
    Right =1378
    Bottom =346
    Left =-1
    Top =-1
    Right =1276
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =148
        Top =6
        Right =244
        Bottom =113
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =5
        Top =9
        Right =101
        Bottom =116
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="Heats"
        Name =""
    End
End
