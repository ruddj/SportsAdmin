dbMemo "SQL" ="SELECT \"D\" AS RecType, Competitors.Surname, Competitors.Gname, Competitors.Sex"
    ", Format(Competitors.DOB,\"mm/dd/yy\") AS Expr1, DLookUp(\"[Mcode]\",\"Miscellan"
    "eous\") AS Expr2, DLookUp(\"[Mteam]\",\"Miscellaneous\") AS Expr3, EventType.Mev"
    "ent, Replace(Nz(CompEvents.Result,\"\"),\"'\",\":\") AS Result, CompEvents.nResu"
    "lt, MeetManagerDivisions.Mdiv\015\012FROM EventType RIGHT JOIN (Competitors LEFT"
    " JOIN ((Events RIGHT JOIN CompEvents ON Events.E_Code = CompEvents.E_Code) LEFT "
    "JOIN MeetManagerDivisions ON Events.Age = MeetManagerDivisions.Eage) ON Competit"
    "ors.PIN = CompEvents.PIN) ON EventType.ET_Code = Events.ET_Code\015\012WHERE ((("
    "EventType.Mevent)<>\"\") AND ((Competitors.Gname)<>\"Team\") AND ((CompEvents.Pl"
    "ace)<=DLookUp(\"[Mtop]\",\"Miscellaneous\")) AND ((CompEvents.F_Lev)=0) AND ((Ev"
    "entType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.Include)=True))\015"
    "\012ORDER BY Competitors.Age DESC , Competitors.Surname, Competitors.Gname;\015\012"
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
        dbText "Name" ="Competitors.Surname"
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
