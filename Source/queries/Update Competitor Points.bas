﻿Operation =4
Option =8
Begin InputTables
    Name ="CompEvents"
    Name ="Heats"
End
Begin OutputColumns
    Name ="CompEvents.Points"
    Expression ="DeterminePoints([Place],[PtScale])"
End
Begin Joins
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
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Heats"
        Name =""
    End
End
