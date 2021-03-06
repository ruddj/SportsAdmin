﻿Operation =1
Option =0
Where ="(((House.Include)=Yes) And ((CompetitorsOrdered.Sex)=\"M\") And ((Val(Competitor"
    "sOrdered.Age))=13))"
Begin InputTables
    Name ="House"
    Name ="CompetitorsOrdered"
End
Begin OutputColumns
    Alias ="fName"
    Expression ="UCase(Trim([Surname]))+\",\"+Trim([Gname])"
    Expression ="CompetitorsOrdered.H_Code"
    Expression ="CompetitorsOrdered.PIN"
    Expression ="House.Include"
End
Begin Joins
    LeftTable ="House"
    RightTable ="CompetitorsOrdered"
    Expression ="House.H_Code = CompetitorsOrdered.H_Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="2"
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
    Left =84
    Top =40
    Right =1002
    Bottom =342
    Left =-1
    Top =-1
    Right =900
    Bottom =127
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="House"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="CompetitorsOrdered"
        Name =""
    End
End
