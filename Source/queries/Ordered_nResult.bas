﻿Operation =1
Option =8
Where ="(((CompEvents.E_Code)=forms!EnterCompetitors!E_Code) And ((CompEvents.Heat)=form"
    "s!EnterCompetitors!Heat) And ((CompEvents.nResult)<>0))"
Begin InputTables
    Name ="CompEvents"
End
Begin OutputColumns
    Expression ="CompEvents.E_Code"
    Expression ="CompEvents.Heat"
    Expression ="CompEvents.nResult"
    Expression ="CompEvents.Place"
End
Begin OrderBy
    Expression ="CompEvents.E_Code"
    Flag =0
    Expression ="CompEvents.Heat"
    Flag =0
    Expression ="CompEvents.nResult"
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
        Name ="CompEvents"
        Name =""
    End
End
