﻿Operation =1
Option =8
Where ="(((Competitors.PIN) Like forms!Compet_search!sPIN) And ((Competitors.Surname) Li"
    "ke forms!Compet_search!sSname) And ((Competitors.Gname) Like forms!Compet_search"
    "!sGname) And ((Competitors.Sex) Like forms!Compet_search!sSex) And ((Competitors"
    ".H_Code) Like forms!Compet_search!sHouse) And ((Competitors.Age) Like forms!Comp"
    "et_search!sAge) And ((Competitors.DOB) Like forms!Compet_search!sDOB))"
Begin InputTables
    Name ="Competitors"
End
Begin OutputColumns
    Expression ="Competitors.PIN"
    Expression ="Competitors.Include"
    Expression ="Competitors.Surname"
    Expression ="Competitors.Gname"
    Expression ="Competitors.Sex"
    Expression ="Competitors.H_Code"
    Expression ="Competitors.Age"
    Expression ="Competitors.DOB"
    Expression ="Competitors.TotPts"
    Expression ="Competitors.Comments"
    Expression ="Competitors.Address1"
    Expression ="Competitors.Address2"
    Expression ="Competitors.Suburb"
    Expression ="Competitors.State"
    Expression ="Competitors.Postcode"
    Expression ="Competitors.Hphone"
    Expression ="Competitors.Wphone"
End
Begin OrderBy
    Expression ="Competitors.Surname"
    Flag =0
    Expression ="Competitors.Gname"
    Flag =0
    Expression ="Competitors.Sex"
    Flag =0
    Expression ="Competitors.Age"
    Flag =0
    Expression ="Competitors.DOB"
    Flag =0
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
        Name ="Competitors"
        Name =""
    End
End
