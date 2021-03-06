﻿Operation =6
Option =8
Begin InputTables
    Name ="Heats"
    Name ="Events"
    Name ="EventType"
    Name ="CompEventsHeats"
End
Begin OutputColumns
    Expression ="Heats.E_Number"
    GroupLevel =2
    Expression ="CompEventsHeats.H_Code"
    GroupLevel =1
    Alias ="SumOfPoints"
    Expression ="Sum(CompEventsHeats.Points)"
End
Begin Joins
    LeftTable ="Heats"
    RightTable ="CompEventsHeats"
    Expression ="Heats.E_Number = CompEventsHeats.E_Number"
    Flag =2
    LeftTable ="Events"
    RightTable ="Heats"
    Expression ="Events.E_Code = Heats.E_Code"
    Flag =3
    LeftTable ="EventType"
    RightTable ="Events"
    Expression ="EventType.ET_Code = Events.ET_Code"
    Flag =3
End
Begin OrderBy
    Expression ="Heats.E_Number"
    Flag =0
End
Begin Groups
    Expression ="Heats.E_Number"
    GroupLevel =2
    Expression ="CompEventsHeats.H_Code"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
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
    ColumnsShown =559
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
        Name ="Events"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="EventType"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="CompEventsHeats"
        Name =""
    End
End
