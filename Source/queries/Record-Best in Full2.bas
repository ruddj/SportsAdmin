﻿Operation =1
Option =8
Where ="(((Better([nResult],[E_Code]))<>False))"
Begin InputTables
    Name ="Records"
End
Begin OutputColumns
    Expression ="Records.E_Code"
    Expression ="Records.nResult"
    Expression ="Records.H_Code"
    Expression ="Records.Surname"
    Expression ="Records.Gname"
    Expression ="Records.Result"
    Expression ="Records.Date"
End
Begin Groups
    Expression ="Records.E_Code"
    GroupLevel =0
    Expression ="Records.nResult"
    GroupLevel =0
    Expression ="Records.H_Code"
    GroupLevel =0
    Expression ="Records.Surname"
    GroupLevel =0
    Expression ="Records.Gname"
    GroupLevel =0
    Expression ="Records.Result"
    GroupLevel =0
    Expression ="Records.Date"
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
        dbText "Name" ="Records.E_Code"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.nResult"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.H_Code"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.Surname"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.Gname"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.Result"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Records.Date"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =93
    Top =318
    Right =907
    Bottom =623
    Left =-1
    Top =-1
    Right =796
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="Records"
        Name =""
    End
End
