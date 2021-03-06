﻿Operation =1
Option =0
Where ="(((House.Flag)=True) AND (([Points-House-Cumulative].Flag)=True))"
Begin InputTables
    Name ="House"
    Name ="Points-House-Cumulative"
End
Begin OutputColumns
    Expression ="House.H_NAme"
    Expression ="[Points-House-Cumulative].E_Number"
    Expression ="[Points-House-Cumulative].Points"
End
Begin Joins
    LeftTable ="Points-House-Cumulative"
    RightTable ="House"
    Expression ="[Points-House-Cumulative].H_ID = House.H_ID"
    Flag =1
End
Begin OrderBy
    Expression ="House.H_NAme"
    Flag =0
    Expression ="[Points-House-Cumulative].E_Number"
    Flag =0
End
Begin Groups
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="[Points-House-Cumulative].E_Number"
    GroupLevel =0
    Expression ="[Points-House-Cumulative].Points"
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
    Begin
        dbText "Name" ="House.H_NAme"
        dbInteger "ColumnWidth" ="3045"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =93
    Top =14
    Right =778
    Bottom =316
    Left =-1
    Top =-1
    Right =667
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =389
        Top =8
        Right =485
        Bottom =115
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =195
        Top =16
        Right =291
        Bottom =108
        Top =0
        Name ="Points-House-Cumulative"
        Name =""
    End
End
