﻿Operation =4
Option =8
Begin InputTables
    Name ="Competitors"
End
Begin OutputColumns
    Name ="Competitors.Gname"
    Expression ="FormatGname([Gname])"
    Name ="Competitors.Surname"
    Expression ="UCase([Surname])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =0
    Top =40
    Right =927
    Bottom =710
    Left =-1
    Top =-1
    Right =909
    Bottom =393
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Competitors"
        Name =""
    End
End
