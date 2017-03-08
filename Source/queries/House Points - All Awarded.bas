dbMemo "SQL" ="SELECT sum([House Points - Total].SumOfPoints + [House Points-Extra Total].SumOf"
    "NumPts) AS AllPoints\015\012FROM [House Points - Total] INNER JOIN [House Points"
    "-Extra Total] ON [House Points - Total].H_Code = [House Points-Extra Total].H_Co"
    "de;\015\012"
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
        dbText "Name" ="AllPoints"
        dbLong "AggregateType" ="-1"
    End
End
