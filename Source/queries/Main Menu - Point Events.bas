dbMemo "SQL" ="TRANSFORM Sum(CompEvents.Points) AS SumOfPoints\015\012SELECT House.H_Code\015\012"
    "FROM House INNER JOIN (EventType INNER JOIN (Competitors INNER JOIN (Events INNE"
    "R JOIN CompEvents ON Events.E_Code = CompEvents.E_Code) ON Competitors.PIN = Com"
    "pEvents.PIN) ON EventType.ET_Code = Events.ET_Code) ON House.H_Code = Competitor"
    "s.H_Code\015\012WHERE (((EventType.Include)=True) AND ((Events.Include)=True) AN"
    "D ((House.Include)=True))\015\012GROUP BY Events.Include, House.Include, House.H"
    "_Code\015\012PIVOT EventType.ET_Des;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Field Shot Put"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track 400M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="House.H_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field Discus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track 100M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field High Jump"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track 1500M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field Javelin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field Long Jump"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Field Triple Jump"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track 200M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track 800M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track House Relay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track Hurdles 100M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track Hurdles 110M"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Track Hurdles 90M"
        dbLong "AggregateType" ="-1"
    End
End
