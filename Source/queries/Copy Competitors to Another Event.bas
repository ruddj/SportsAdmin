﻿Operation =3
Name ="CompEvents"
Option =8
Begin InputTables
    Name ="CompEvents"
    Name ="Events"
End
Begin OutputColumns
    Name ="PIN"
    Expression ="CompEvents.PIN"
    Alias ="NewEcode"
    Name ="E_Code"
    Expression ="DLookUp(\"[E_Code]\",\"Events in Full\",\"[ET_Code]=1 and [Heat]=\" & [Heat] & \""
        " and [F_Lev]=\" & [F_Lev] & \" and [Sex]=\"\"\" & [Events].[Sex] & \"\"\" and [A"
        "ge]=\"\"\" & [Events].[Age] & \"\"\"\")"
    Name ="Place"
    Expression ="CompEvents.Place"
    Name ="TTres"
    Expression ="CompEvents.TTres"
    Name ="Heat"
    Expression ="CompEvents.Heat"
    Name ="Lane"
    Expression ="CompEvents.Lane"
    Name ="Result"
    Expression ="CompEvents.Result"
    Name ="nResult"
    Expression ="CompEvents.nResult"
    Name ="F_Lev"
    Expression ="CompEvents.F_Lev"
    Name ="Memo"
    Expression ="CompEvents.Memo"
    Name ="Points"
    Expression ="CompEvents.Points"
End
Begin Joins
    LeftTable ="CompEvents"
    RightTable ="Events"
    Expression ="CompEvents.E_Code = Events.E_Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CompEvents.Place"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Heat"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Lane"
        dbInteger "ColumnWidth" ="585"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Result"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.nResult"
        dbInteger "ColumnWidth" ="705"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.F_Lev"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.Memo"
        dbInteger "ColumnWidth" ="570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewEcode"
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
    ColumnsShown =651
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
        Name ="Events"
        Name =""
    End
End
