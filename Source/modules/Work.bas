Option Compare Database
Option Explicit

Private Function FixCompEvents()

  Dim Ers As Recordset
  Dim RS As Recordset
  Dim MoveCompetitor  As Boolean
  Dim A As String
  
  Set Ers = CurrentDb.OpenRecordset("Events")
  Q = "SELECT CompEvents.PIN, Events.ET_Code, CompEvents.E_Code, CompEvents.Heat, CompEvents.F_Lev, Competitors.Age AS CompetitorAge, Events.Age AS EventAge, Events.Sex"
  Q = Q & " FROM (Events INNER JOIN Heats ON Events.E_Code = Heats.E_Code) INNER JOIN (Competitors INNER JOIN CompEvents ON Competitors.PIN = CompEvents.PIN) ON (Heats.Heat = CompEvents.Heat) AND (Heats.F_Lev = CompEvents.F_Lev) AND (Heats.E_Code = CompEvents.E_Code);"

  Set RS = CurrentDb.OpenRecordset(Q)
  
  Do Until RS.EOF
    MoveCompetitor = False
    If Right(RS!EventAge, 2) = "_O" Then
      If RS!CompetitorAge < Val(RS!EventAge) Then
        MoveCompetitor = True
      End If
    
    ElseIf Right(RS!EventAge, 2) = "_U" Then
      If RS!CompetitorAge > Val(RS!EventAge) Then
        MoveCompetitor = True
      End If
    
    ElseIf Val(RS!EventAge) <> RS!CompetitorAge Then
        MoveCompetitor = True
    End If
    
    If MoveCompetitor Then
      Select Case RS!CompetitorAge
        Case 12: A = "13_U"
        Case 13: A = "13_U"
        Case 14: A = "14"
        Case 15: A = "15"
        Case 16: A = "16"
        Case 17: A = "17_O"
        Case 18: A = "18_O"
        Case Else: Stop
      End Select
      
      Ers.FindFirst "[ET_Code]=" & RS![ET_Code] & " AND [Sex]='" & RS!Sex & "' AND [Age]='" & A & "'"
      If Ers.NoMatch Then
        Debug.Print RS![ET_Code]
        'Stop
      Else
        Debug.Print "Changed"
        RS.Edit
        RS!E_Code = Ers!E_Code
        RS.Update
      End If
    End If
  
    RS.MoveNext
  Loop
End Function