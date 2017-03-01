Operation =1
Option =8
Begin InputTables
    Name ="Heats"
    Name ="House"
    Name ="CompEvents"
    Name ="Competitors"
    Name ="PointsScale"
End
Begin OutputColumns
    Expression ="House.H_NAme"
    Expression ="House.CompPool"
    Alias ="SumOfPoints"
    Expression ="Sum(PointsScale.Points)"
End
Begin Joins
    LeftTable ="Heats"
    RightTable ="PointsScale"
    Expression ="Heats.PtScale = PointsScale.PtScale"
    Flag =1
    LeftTable ="CompEvents"
    RightTable ="PointsScale"
    Expression ="CompEvents.Place = PointsScale.Place"
    Flag =1
    LeftTable ="Competitors"
    RightTable ="CompEvents"
    Expression ="Competitors.PIN = CompEvents.PIN"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.E_Code = CompEvents.E_Code"
    Flag =1
    LeftTable ="Heats"
    RightTable ="CompEvents"
    Expression ="Heats.Heat = CompEvents.Heat"
    Flag =1
    LeftTable ="House"
    RightTable ="Competitors"
    Expression ="House.H_Code = Competitors.H_Code"
    Flag =1
End
Begin Groups
    Expression ="House.H_NAme"
    GroupLevel =0
    Expression ="House.CompPool"
    GroupLevel =0
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
    Begin
        dbText "Name" ="SumOfPoints"
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
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Heats"
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
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="PointsScale"
        Name =""
    End
End
