﻿Operation =1
Option =0
Begin InputTables
    Name ="CompEvents"
    Name ="Heats"
End
Begin OutputColumns
    Expression ="CompEvents.PIN"
    Expression ="CompEvents.E_Code"
    Expression ="CompEvents.Place"
    Expression ="CompEvents.TTres"
    Expression ="CompEvents.Heat"
    Expression ="CompEvents.Result"
    Expression ="CompEvents.nResult"
    Expression ="CompEvents.F_Lev"
    Expression ="Heats.EffectsRecords"
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
    Left =84
    Top =40
    Right =1378
    Bottom =517
    Left =-1
    Top =-1
    Right =1276
    Bottom =291
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =246
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =482
        Bottom =243
        Top =0
        Name ="Heats"
        Name =""
    End
End
