﻿Operation =1
Option =0
Where ="(((Competitors.Sex)=forms!EnterCompetitors!SexFld) And ((Competitors.Age)=forms!"
    "EnterCompetitors!AgeFld))"
Begin InputTables
    Name ="Competitors"
End
Begin OutputColumns
    Alias ="fName"
    Expression ="UCase(Trim([Surname]))+\", \"+Trim([Gname])"
    Expression ="Competitors.H_Code"
    Expression ="Competitors.PIN"
End
Begin OrderBy
    Expression ="UCase(Trim([Surname]))+\", \"+Trim([Gname])"
    Flag =0
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
        dbText "Name" ="fName"
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
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Competitors"
        Name =""
    End
End
