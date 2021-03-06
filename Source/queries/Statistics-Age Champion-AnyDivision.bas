﻿dbMemo "SQL" ="SELECT DISTINCTROW Trim(UCase([Surname])) & \", \" & [Gname] & \" (\" & [Competi"
    "tors].[Age] & \")\" AS Fullname, [AgeOpen].[Eage] & ' ' & [Sex Sub].[Sex Sub] AS"
    " AgeSex, House.H_NAme, Sum(CompEvents.Points) AS SumOfPoints, Competitors.Age, E"
    "ventType.Flag, House.Flag\015\012FROM AgeOpen RIGHT JOIN (House RIGHT JOIN ((Eve"
    "ntType RIGHT JOIN (([Sex Sub] RIGHT JOIN Events ON [Sex Sub].Sex = Events.Sex) R"
    "IGHT JOIN Heats ON Events.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.E"
    "T_Code) RIGHT JOIN (Competitors RIGHT JOIN CompEvents ON Competitors.PIN = CompE"
    "vents.PIN) ON (Heats.Heat = CompEvents.Heat) AND (Heats.F_Lev = CompEvents.F_Lev"
    ") AND (Heats.E_Code = CompEvents.E_Code)) ON House.H_Code = Competitors.H_Code) "
    "ON AgeOpen.Cage = Competitors.Age\015\012GROUP BY Trim(UCase([Surname])) & \", \""
    " & [Gname] & \" (\" & [Competitors].[Age] & \")\", [AgeOpen].[Eage] & ' ' & [Sex"
    " Sub].[Sex Sub], House.H_NAme, Competitors.Age, EventType.Flag, House.Flag, UCas"
    "e([Gname])\015\012HAVING (((Competitors.Age) Is Not Null) AND ((EventType.Flag)="
    "Yes) AND ((House.Flag)=Yes) AND ((UCase([Gname]))<>\"TEAM\"))\015\012ORDER BY [A"
    "geOpen].[Eage] & ' ' & [Sex Sub].[Sex Sub], Sum(CompEvents.Points) DESC;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="Fullname"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.Flag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.Flag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgeSex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.H_NAme"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPoints"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.Age"
        dbLong "AggregateType" ="-1"
    End
End
