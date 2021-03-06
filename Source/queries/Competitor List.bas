﻿Operation =1
Option =8
Where ="(((House.Flag)=Yes) AND ((Competitors.Include)=Yes))"
Begin InputTables
    Name ="House"
    Name ="Competitors"
    Name ="Sex Sub"
    Name ="CompetitorEventAge"
End
Begin OutputColumns
    Alias ="FullName"
    Expression ="UCase([Surname]) & \", \" & [Gname]"
    Expression ="Competitors.Sex"
    Expression ="[Sex Sub].[Sex Sub]"
    Expression ="Competitors.DOB"
    Expression ="Competitors.H_Code"
    Expression ="Competitors.Age"
    Alias ="AAge"
    Expression ="CompetitorEventAge.Eage"
    Expression ="House.H_NAme"
    Expression ="House.Flag"
    Expression ="Competitors.Include"
    Alias ="nAge"
    Expression ="Val([Age])"
End
Begin Joins
    LeftTable ="Competitors"
    RightTable ="Sex Sub"
    Expression ="Competitors.Sex = [Sex Sub].Sex"
    Flag =1
    LeftTable ="Competitors"
    RightTable ="CompetitorEventAge"
    Expression ="Competitors.Age = CompetitorEventAge.Cage"
    Flag =1
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =1
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
        dbText "Name" ="FullName"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.Flag"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AAge"
    End
    Begin
        dbText "Name" ="nAge"
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
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =83
        Top =0
        Name ="Sex Sub"
        Name =""
    End
    Begin
        Left =451
        Top =44
        Right =645
        Bottom =121
        Top =0
        Name ="CompetitorEventAge"
        Name =""
    End
End
