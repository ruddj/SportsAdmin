﻿Operation =3
Name ="Competitors"
Option =8
Begin InputTables
    Name ="TEMP1"
End
Begin OutputColumns
    Name ="Gname"
    Expression ="TEMP1.Given"
    Name ="Surname"
    Expression ="TEMP1.Surname"
    Name ="DOB"
    Expression ="TEMP1.DOB"
    Name ="Sex"
    Expression ="TEMP1.Sex"
    Name ="H_Code"
    Expression ="TEMP1.House"
    Name ="Age"
    Expression ="TEMP1.Age"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
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
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="TEMP1"
        Name =""
    End
End
