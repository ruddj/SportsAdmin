﻿Operation =1
Option =1
Where ="(((EnterCompetitors.ET_Code)=Forms!Reports_Event!Event_DD) And ((EnterCompetitor"
    "s.Age) Like Val(Forms!Reports_Event!Age_EB) Or (EnterCompetitors.Age) Like Forms"
    "!Reports_Event!Age_EB) And ((EnterCompetitors.Heat) Like Val(Forms!Reports_Event"
    "!Heat_EB) Or (EnterCompetitors.Heat) Like Forms!Reports_Event!Heat_EB) And ((Ent"
    "erCompetitors.Sex) Like Forms!Reports_Event!Sex_DD))"
Begin InputTables
    Name ="EnterCompetitors"
End
Begin OutputColumns
    Expression ="EnterCompetitors.ET_Code"
    Expression ="EnterCompetitors.Age"
    Expression ="EnterCompetitors.Heat"
    Expression ="EnterCompetitors.Sex"
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
        Name ="EnterCompetitors"
        Name =""
    End
End
