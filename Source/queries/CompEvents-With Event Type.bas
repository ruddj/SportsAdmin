﻿Operation =1
Option =0
Begin InputTables
    Name ="CompEvents"
    Name ="Events"
    Name ="EventType"
End
Begin OutputColumns
    Expression ="CompEvents.*"
    Expression ="EventType.ET_Des"
    Expression ="Events.Sex"
    Expression ="Events.Age"
    Expression ="EventType.ET_Code"
End
Begin Joins
    LeftTable ="CompEvents"
    RightTable ="Events"
    Expression ="CompEvents.E_Code = Events.E_Code"
    Flag =1
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
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
    Top =14
    Right =778
    Bottom =319
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
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="Events"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =113
        Top =0
        Name ="EventType"
        Name =""
    End
End
