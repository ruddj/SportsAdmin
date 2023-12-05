Operation =1
Option =0
Begin InputTables
    Name ="Heats"
    Name ="Competitors"
    Name ="CompEvents"
    Name ="House"
End
Begin OutputColumns
    Expression ="Competitors.H_Code"
    Expression ="House.H_ID"
    Expression ="Heats.HE_Code"
    Expression ="CompEvents.PIN"
    Expression ="CompEvents.Place"
    Expression ="CompEvents.TTres"
    Expression ="CompEvents.Heat"
    Alias ="join"
    Expression ="Str([CompEvents].[E_Code])+Str([CompEvents].[Heat])"
    Expression ="CompEvents.nResult"
    Expression ="CompEvents.Result"
    Expression ="CompEvents.E_Code"
    Expression ="CompEvents.F_Lev"
    Expression ="CompEvents.Lane"
    Expression ="Heats.AllNames"
    Expression ="CompEvents.Memo"
    Expression ="CompEvents.Points"
End
Begin Joins
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
    Expression ="Heats.F_Lev = CompEvents.F_Lev"
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
Begin OrderBy
    Expression ="CompEvents.Lane"
    Flag =0
    Expression ="Competitors.Surname"
    Flag =0
    Expression ="Competitors.Gname"
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
    Begin
        dbText "Name" ="join"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.H_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =10
    Right =1180
    Bottom =477
    Left =-1
    Top =-1
    Right =1050
    Bottom =311
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =312
        Top =0
        Name ="Heats"
        Name =""
    End
    Begin
        Left =396
        Top =12
        Right =492
        Bottom =285
        Top =0
        Name ="Competitors"
        Name =""
    End
    Begin
        Left =236
        Top =7
        Right =332
        Bottom =286
        Top =0
        Name ="CompEvents"
        Name =""
    End
    Begin
        Left =574
        Top =0
        Right =670
        Bottom =258
        Top =0
        Name ="House"
        Name =""
    End
End
