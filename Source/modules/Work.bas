Option Compare Database
Option Explicit

Private Function FixCompEvents()

  Dim Ers As Recordset
  Dim Rs As Recordset
  Dim MoveCompetitor  As Boolean
  Dim A As String, Q As String
  
  Set Ers = CurrentDb.OpenRecordset("Events")
  Q = "SELECT CompEvents.PIN, Events.ET_Code, CompEvents.E_Code, CompEvents.Heat, CompEvents.F_Lev, Competitors.Age AS CompetitorAge, Events.Age AS EventAge, Events.Sex"
  Q = Q & " FROM (Events INNER JOIN Heats ON Events.E_Code = Heats.E_Code) INNER JOIN (Competitors INNER JOIN CompEvents ON Competitors.PIN = CompEvents.PIN) ON (Heats.Heat = CompEvents.Heat) AND (Heats.F_Lev = CompEvents.F_Lev) AND (Heats.E_Code = CompEvents.E_Code);"

  Set Rs = CurrentDb.OpenRecordset(Q)
  
  Do Until Rs.EOF
    MoveCompetitor = False
    If Right(Rs!EventAge, 2) = "_O" Then
      If Rs!CompetitorAge < Val(Rs!EventAge) Then
        MoveCompetitor = True
      End If
    
    ElseIf Right(Rs!EventAge, 2) = "_U" Then
      If Rs!CompetitorAge > Val(Rs!EventAge) Then
        MoveCompetitor = True
      End If
    
    ElseIf Val(Rs!EventAge) <> Rs!CompetitorAge Then
        MoveCompetitor = True
    End If
    
    If MoveCompetitor Then
      Select Case Rs!CompetitorAge
        Case 12: A = "13_U"
        Case 13: A = "13_U"
        Case 14: A = "14"
        Case 15: A = "15"
        Case 16: A = "16"
        Case 17: A = "17_O"
        Case 18: A = "18_O"
        Case Else: Stop
      End Select
      
      Ers.FindFirst "[ET_Code]=" & Rs![ET_Code] & " AND [Sex]='" & Rs!Sex & "' AND [Age]='" & A & "'"
      If Ers.NoMatch Then
        Debug.Print Rs![ET_Code]
        'Stop
      Else
        Debug.Print "Changed"
        Rs.Edit
        Rs!E_Code = Ers!E_Code
        Rs.Update
      End If
    End If
  
    Rs.MoveNext
  Loop
End Function