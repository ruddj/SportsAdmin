﻿Operation =1
Option =8
Begin InputTables
    Name ="Records"
    Name ="Events"
    Name ="EventType"
    Name ="Units"
End
Begin OutputColumns
    Alias ="BestResult"
    Expression ="IIf([Order]=\"ASC\",[nResult],(1/[nResult]))"
    Expression ="Records.E_Code"
    Expression ="Records.Surname"
    Expression ="Records.Gname"
    Expression ="Records.H_Code"
    Expression ="Records.Date"
    Expression ="Records.Comments"
    Expression ="Records.nResult"
    Expression ="Records.Result"
End
Begin Joins
    LeftTable ="EventType"
    RightTable ="Units"
    Expression ="EventType.Units = Units.DisplayUnit"
    Flag =1
    LeftTable ="Events"
    RightTable ="Records"
    Expression ="Events.E_Code = Records.E_Code"
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
    Begin
        dbText "Name" ="BestResult"
    End
End
Begin
    State =0
    Left =84
    Top =14
    Right =1002
    Bottom =319
    Left =-1
    Top =-1
    Right =900
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
        Name ="Records"
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
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =113
        Top =0
        Name ="Units"
        Name =""
    End
End
