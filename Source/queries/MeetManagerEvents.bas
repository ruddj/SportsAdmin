﻿dbMemo "SQL" ="SELECT \"D;\" & Competitors.Surname & \";\" & Competitors.Gname & \";;\" & Compe"
    "titors.Sex & \";\" & Format(Competitors.DOB,\"mm/dd/yy\") & \";\" & DLookUp(\"[M"
    "code]\",\"Miscellaneous\") & \";\" & DLookUp(\"[Mteam]\",\"Miscellaneous\") & \""
    ";;;\" & EventType.Mevent & \";\" & Replace(Nz(CompEvents.Result,\"\") , \"'\",\""
    ":\") & \";M;\" & MeetManagerDivisions.Mdiv & \";\" AS EntryRecord\015\012FROM (E"
    "ventType RIGHT JOIN (Competitors LEFT JOIN (Events RIGHT JOIN CompEvents ON Even"
    "ts.E_Code = CompEvents.E_Code) ON Competitors.PIN = CompEvents.PIN) ON EventType"
    ".ET_Code = Events.ET_Code) LEFT JOIN MeetManagerDivisions ON Events.Age = MeetMa"
    "nagerDivisions.Eage\015\012WHERE (((Competitors.Gname)<>\"Team\") AND ((CompEven"
    "ts.Place)<=DLookUp(\"[Mtop]\",\"Miscellaneous\")) AND ((CompEvents.F_Lev)=0) AND"
    " ((EventType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.Include)=Tr"
    "ue) AND ((EventType.Mevent)<>\"\"))\015\012ORDER BY Competitors.Age DESC , Compe"
    "titors.Surname, Competitors.Gname;\015\012"
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
        dbText "Name" ="EntryRecord"
        dbInteger "ColumnWidth" ="7035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
