﻿Operation =1
Option =8
Begin InputTables
    Name ="House Points-Extra"
    Name ="House"
End
Begin OutputColumns
    Expression ="House.H_Code"
    Expression ="House.H_NAme"
    Expression ="[House Points-Extra].H_ID"
    Alias ="SumOfNumPts"
    Expression ="Sum([House Points-Extra].NumPts)"
End
Begin Joins
    LeftTable ="House"
    RightTable ="House Points-Extra"
    Expression ="House.H_ID = [House Points-Extra].H_ID"
    Flag =2
End
Begin Groups
    Expression ="House.H_Code"
    GroupLevel =0
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="[House Points-Extra].H_ID"
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
        dbText "Name" ="SumOfNumPts"
        dbLong "AggregateType" ="-1"
    End
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
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="House Points-Extra"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="House"
        Name =""
    End
End
