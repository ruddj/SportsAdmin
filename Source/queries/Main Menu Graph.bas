﻿Operation =1
Option =8
Begin InputTables
    Name ="House Points - Grand Total (All Events)"
End
Begin OutputColumns
    Expression ="[House Points - Grand Total (All Events)].H_Code"
    Expression ="[House Points - Grand Total (All Events)].GrandTotal"
End
Begin Groups
    Expression ="[House Points - Grand Total (All Events)].H_Code"
    GroupLevel =0
    Expression ="[House Points - Grand Total (All Events)].GrandTotal"
    GroupLevel =0
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
End
Begin
    State =0
    Left =98
    Top =44
    Right =1029
    Bottom =349
    Left =-1
    Top =-1
    Right =913
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
        Name ="House Points - Grand Total (All Events)"
        Name =""
    End
End
