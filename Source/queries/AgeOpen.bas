﻿dbMemo "SQL" ="SELECT DISTINCT Competitors.Age as Cage, Competitors.Age as Eage\015\012FROM Com"
    "petitors\015\012WHERE (((Competitors.Age) Is Not Null)) AND (((Competitors.Age)<"
    "DLookUp(\"[OpenAge]\",\"Miscellaneous\")))\015\012UNION SELECT DISTINCT Competit"
    "ors.Age AS Cage, 'Open' AS Eage\015\012FROM Competitors\015\012WHERE (((Competit"
    "ors.Age) Is Not Null And (Competitors.Age)>=DLookUp(\"[OpenAge]\",\"Miscellaneou"
    "s\")));\015\012"
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
        dbText "Name" ="Cage"
    End
    Begin
        dbText "Name" ="Eage"
    End
End
