Option Compare Database
Option Explicit

Private Function FixCompEvents()

  Dim Ers As Recordset
  Dim rs As Recordset
  Dim MoveCompetitor  As Boolean
  Dim A As String
  
  Set Ers = CurrentDb.OpenRecordset("Events")
  Q = "SELECT CompEvents.PIN, Events.ET_Code, CompEvents.E_Code, CompEvents.Heat, CompEvents.F_Lev, Competitors.Age AS CompetitorAge, Events.Age AS EventAge, Events.Sex"
  Q = Q & " FROM (Events INNER JOIN Heats ON Events.E_Code = Heats.E_Code) INNER JOIN (Competitors INNER JOIN CompEvents ON Competitors.PIN = CompEvents.PIN) ON (Heats.Heat = CompEvents.Heat) AND (Heats.F_Lev = CompEvents.F_Lev) AND (Heats.E_Code = CompEvents.E_Code);"

  Set rs = CurrentDb.OpenRecordset(Q)
  
  Do Until rs.EOF
    MoveCompetitor = False
    If Right(rs!EventAge, 2) = "_O" Then
      If rs!CompetitorAge < Val(rs!EventAge) Then
        MoveCompetitor = True
      End If
    
    ElseIf Right(rs!EventAge, 2) = "_U" Then
      If rs!CompetitorAge > Val(rs!EventAge) Then
        MoveCompetitor = True
      End If
    
    ElseIf Val(rs!EventAge) <> rs!CompetitorAge Then
        MoveCompetitor = True
    End If
    
    If MoveCompetitor Then
      Select Case rs!CompetitorAge
        Case 12: A = "13_U"
        Case 13: A = "13_U"
        Case 14: A = "14"
        Case 15: A = "15"
        Case 16: A = "16"
        Case 17: A = "17_O"
        Case 18: A = "18_O"
        Case Else: Stop
      End Select
      
      Ers.FindFirst "[ET_Code]=" & rs![ET_Code] & " AND [Sex]='" & rs!Sex & "' AND [Age]='" & A & "'"
      If Ers.NoMatch Then
        Debug.Print rs![ET_Code]
        'Stop
      Else
        Debug.Print "Changed"
        rs.Edit
        rs!E_Code = Ers!E_Code
        rs.Update
      End If
    End If
  
    rs.MoveNext
  Loop
End Function