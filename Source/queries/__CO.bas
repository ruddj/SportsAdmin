﻿Operation =1
Option =0
Where ="((([CompetitorsOrdered].[Flag])=True))"
Begin InputTables
    Name ="CompetitorsOrdered"
End
Begin OutputColumns
    Expression ="CompetitorsOrdered.PIN"
    Expression ="CompetitorsOrdered.Gname"
    Expression ="CompetitorsOrdered.Surname"
    Alias ="Expr1"
    Expression ="CompetitorsOrdered.Flag"
    Alias ="Expr2"
    Expression ="CompetitorsOrdered.Order"
End
Begin OrderBy
    Expression ="CompetitorsOrdered.Order"
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
        dbText "Name" ="Expr1"
    End
    Begin
        dbText "Name" ="Expr2"
    End
End
Begin
    State =0
    Left =84
    Top =14
    Right =778
    Bottom =316
    Left =-1
    Top =-1
    Right =676
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="CompetitorsOrdered"
        Name =""
    End
End