﻿Operation =1
Option =1
Where ="(((EnterCompetitors.ET_Code) Like Forms!EnterCompetitors!Event_DD) And ((EnterCo"
    "mpetitors.Age) Like Val(Forms!EnterCompetitors!Age_EB) Or (EnterCompetitors.Age)"
    " Like Forms!EnterCompetitors!Age_EB) And ((EnterCompetitors.Heat) Like Val(Forms"
    "!EnterCompetitors!Heat_EB) Or (EnterCompetitors.Heat) Like Forms!EnterCompetitor"
    "s!Heat_EB) And ((EnterCompetitors.Sex) Like Forms!EnterCompetitors!Sex_DD) And ("
    "(EnterCompetitors.E_Number) Like Forms!EnterCompetitors!Enum) And ((EnterCompeti"
    "tors.F_Lev) Like Forms!EnterCompetitors!Flevel))"
Begin InputTables
    Name ="EnterCompetitors"
End
Begin OutputColumns
    Expression ="EnterCompetitors.ET_Code"
    Expression ="EnterCompetitors.Age"
    Expression ="EnterCompetitors.Heat"
    Expression ="EnterCompetitors.Sex"
    Expression ="EnterCompetitors.E_Number"
    Expression ="EnterCompetitors.F_Lev"
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
