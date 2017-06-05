dbMemo "SQL" ="SELECT \"D\" AS RecType, Competitors.Surname, Competitors.Gname, Competitors.Sex"
    ", Format(Competitors.DOB,\"mm/dd/yy\"), DLookUp(\"[Mcode]\",\"Miscellaneous\"), "
    "DLookUp(\"[Mteam]\",\"Miscellaneous\"), EventType.Mevent, Replace(CompEvents.Res"
    "ult , \"'\",\":\") AS Result, CompEvents.nResult, MeetManagerDivisions.Mdiv\015\012"
    "FROM (EventType RIGHT JOIN (Competitors LEFT JOIN (Events RIGHT JOIN CompEvents "
    "ON Events.E_Code = CompEvents.E_Code) ON Competitors.PIN = CompEvents.PIN) ON Ev"
    "entType.ET_Code = Events.ET_Code) LEFT JOIN MeetManagerDivisions ON Events.Age ="
    " MeetManagerDivisions.Eage\015\012WHERE (((Competitors.Gname)<>\"Team\") AND ((C"
    "ompEvents.Place)<=DLookUp(\"[Mtop]\",\"Miscellaneous\")) AND ((CompEvents.F_Lev)"
    "=0) AND ((EventType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.Incl"
    "ude)=True) AND ((EventType.Mevent)<>\"\"))\015\012ORDER BY Competitors.Age DESC "
    ", Competitors.Surname, Competitors.Gname;\015\012"
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
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.Surname"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1005"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.Gname"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Competitors.Sex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventType.Mevent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Result"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CompEvents.nResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MeetManagerDivisions.Mdiv"
        dbLong "AggregateType" ="-1"
    End
End
