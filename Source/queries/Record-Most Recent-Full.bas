﻿Operation =1
Option =8
Begin InputTables
    Name ="Records"
    Name ="Record-Most Recent"
End
Begin OutputColumns
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
    LeftTable ="Records"
    RightTable ="Record-Most Recent"
    Expression ="Records.E_Code = [Record-Most Recent].E_Code"
    Flag =1
    LeftTable ="Record-Most Recent"
    RightTable ="Records"
    Expression ="[Record-Most Recent].MaxOfDate = Records.Date"
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
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Records"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Record-Most Recent"
        Name =""
    End
End
