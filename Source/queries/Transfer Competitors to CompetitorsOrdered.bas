﻿Operation =3
Name ="CompetitorsOrdered"
Option =8
Begin InputTables
    Name ="Competitors"
End
Begin OutputColumns
    Name ="PIN"
    Expression ="Competitors.PIN"
    Name ="Include"
    Expression ="Competitors.Include"
    Name ="Surname"
    Expression ="Competitors.Surname"
    Name ="Gname"
    Expression ="Competitors.Gname"
    Name ="Sex"
    Expression ="Competitors.Sex"
    Name ="H_Code"
    Expression ="Competitors.H_Code"
    Name ="H_ID"
    Expression ="Competitors.H_ID"
    Name ="DOB"
    Expression ="Competitors.DOB"
    Name ="TotPts"
    Expression ="Competitors.TotPts"
    Name ="Comments"
    Expression ="Competitors.Comments"
    Name ="Address1"
    Expression ="Competitors.Address1"
    Name ="Address2"
    Expression ="Competitors.Address2"
    Name ="Suburb"
    Expression ="Competitors.Suburb"
    Name ="State"
    Expression ="Competitors.State"
    Name ="Postcode"
    Expression ="Competitors.Postcode"
    Name ="Hphone"
    Expression ="Competitors.Hphone"
    Name ="Wphone"
    Expression ="Competitors.Wphone"
    Name ="Age"
    Expression ="Competitors.Age"
End
Begin OrderBy
    Expression ="Competitors.Surname"
    Flag =0
    Expression ="Competitors.Gname"
    Flag =0
    Expression ="Competitors.Sex"
    Flag =0
    Expression ="Competitors.H_Code"
    Flag =0
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
        Name ="Competitors"
        Name =""
    End
End
