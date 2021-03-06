﻿Operation =3
Name ="Records"
Option =8
Begin InputTables
    Name ="Events"
    Name ="House"
End
Begin OutputColumns
    Name ="nResult"
    Expression ="Events.nRecord"
    Name ="E_Code"
    Expression ="Events.E_Code"
    Name ="H_Code"
    Expression ="House.H_Code"
    Name ="Surname"
    Expression ="Events.RecName"
    Alias ="Expr1"
    Name ="Gname"
    Expression ="\" \""
    Alias ="Expr2"
    Name ="Date"
    Expression ="#1/3/1997#"
    Name ="Result"
    Expression ="Events.Record"
End
Begin Joins
    LeftTable ="Events"
    RightTable ="House"
    Expression ="Events.RecHouse = House.H_ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Expr1"
    End
    Begin
        dbText "Name" ="Expr2"
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
        Name ="Events"
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
