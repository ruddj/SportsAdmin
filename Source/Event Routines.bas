Option Compare Database
Option Explicit

Public Function AutomaticallyCreateHeatsAndFinals(ET_Code As Long, Optional NewEvent As Boolean, _
                                                  Optional Quiet As Boolean = False, Optional ClearExisting As Boolean = True) As Boolean
On Error GoTo AutomaticallyCreateHeatsAndFinals_Err

  Dim Criteria As String
  Dim Eset As Recordset, FLSet As Recordset
  Dim Hset As Recordset, CEset As Recordset
  Dim Count As Long, Q As String, msg As String, q1 As String, i As Integer
  Dim Response  As Variant, CountRecs As Long
  Dim ReturnValue As Variant, First_FL As Boolean, AddHeat As Boolean
  Dim Ecode  As Long, ET_Des As Variant, NoHeats  As Integer
  
  'ETcode = Forms![EventType]![ET_Code]

  Q = "SELECT DISTINCTROW Events.ET_Code, Events.E_Code FROM Events "
  Q = Q & "WHERE (Events.ET_Code = " & ET_Code & ")"
    
  Set Eset = CurrentDb.OpenRecordset(Q, DB_OPEN_DYNASET)   ' Create dynaset.

  If DCount("[E_Code]", "Events", "[ET_Code]=" & ET_Code) = 0 Then
     Response = MsgBox("There are no Divisions set up for this event.  Set up the Divisions before attempting to create Heats and Final Levels.", vbOKOnly + vbInformation)
     AutomaticallyCreateHeatsAndFinals = False
  Else
    ET_Des = DLookup("[ET_Des]", "EventType", "[ET_code]=" & ET_Code)
    If IsNull(ET_Des) Then
      MsgBox "An unexpected error has occured: the EventType is missing", vbExclamation
      AutomaticallyCreateHeatsAndFinals = False
      GoTo AutomaticallyCreateHeatsAndFinals_Exit
    End If
    
    If NewEvent = True Then
      Response = vbYes
    Else
      If ClearExisting Then
        msg = "This action will remove all competitors from the "
        msg = msg & ET_Des & ".  Do you want to continue?"
        
        If Not Quiet Then
          Response = MsgBox(msg, vbYesNo + vbDefaultButton2 + vbExclamation, "Confirm Creation of Heats and Finals")
        Else
          Response = vbYes
        End If
      Else
        Response = vbYes
      End If
    End If
    If Response = vbYes Then
      PleaseWaitMsg = "Creating heats and finals for all divisions of the " & ET_Des & ".  Please wait ..."
      DoCmd.RunMacro "ShowPleaseWait"

      Q = "SELECT DISTINCTROW Final_Lev.ET_Code, Final_Lev.F_Lev, Final_Lev.NoHeats, Final_Lev.PtScale, Final_Lev.ProType, Final_Lev.UseTimes, Final_Lev.EffectsRecords "
      Q = Q & "FROM Final_Lev "
      Q = Q & "WHERE (Final_Lev.ET_Code = " & ET_Code & ") "
      Q = Q & "ORDER BY F_Lev Desc"
      
      Set FLSet = CurrentDb.OpenRecordset(Q, DB_OPEN_DYNASET)   ' Create dynaset.
      If FLSet.BOF Then
        Response = MsgBox("The final and heat information has not been provided so none can be created.", vbOKOnly + vbInformation)
        AutomaticallyCreateHeatsAndFinals = False
        GoTo AutomaticallyCreateHeatsAndFinals_Exit
      End If
            
      Eset.MoveLast
      CountRecs = Eset.RecordCount

      FLSet.MoveLast
      CountRecs = CountRecs * FLSet.RecordCount

      msg = "Creating Heats and Finals ..."
      ReturnValue = SysCmd(SYSCMD_INITMETER, msg, CountRecs)    ' Display message in status bar.
      

      ' ****************************************************************
      ' Delete all Heats and CompEvents with selected ET_Code

      Eset.MoveFirst
      Count = 1

      While Not Eset.EOF
          
          SysCmd SYSCMD_UPDATEMETER, Count   ' Update meter.
          
          Ecode = Eset!E_Code
          
          If ClearExisting Then
            q1 = "Delete * from CompEvents where CompEvents.E_Code = " & Ecode
            'DoCmd RunSQL q1
            CurrentDb.Execute (q1)
  
            q1 = "Delete * from Heats where Heats.E_Code = " & Ecode
            'DoCmd RunSQL q1
            CurrentDb.Execute (q1)
          
          End If

          FLSet.MoveFirst
          
          Set Hset = CurrentDb.OpenRecordset("Heats", DB_OPEN_DYNASET)   ' Create dynaset.
                  
          First_FL = True
          While Not FLSet.EOF
                  
            NoHeats = FLSet!NoHeats
            If NoHeats > 999 Then
              NoHeats = 999
            ElseIf NoHeats < 1 Then
              NoHeats = 1
            End If

            Count = Count + 1
            
            For i = 1 To NoHeats
              If Not ClearExisting Then
                Hset.FindFirst "[E_Code]=" & Ecode & " AND [F_Lev]=" & FLSet!F_Lev & " AND [Heat]=" & i
                If Hset.NoMatch Then
                  AddHeat = True
                Else
                  AddHeat = False
                End If
              Else
                AddHeat = True
              End If
              
              If AddHeat Then
                'Hset.Edit
                Hset.AddNew
                Hset!E_Code = Ecode
                Hset!Heat = i
                Hset!PtScale = FLSet!PtScale
                Hset!E_Number = 0
                Hset!E_Time = Null
                Hset!F_Lev = FLSet!F_Lev
                Hset!Pro_Type = FLSet!ProType
                Hset!UseTimes = FLSet!UseTimes
                Hset!EffectsRecords = FLSet!EffectsRecords
                If First_FL Then
                    Hset!Status = 1
                Else
                    Hset!Status = 0
                End If
                
                Hset.Update
              End If
            Next i

            First_FL = False
            FLSet.MoveNext
          Wend

          Hset.Close
          Eset.MoveNext
      Wend

      Eset.Close
      FLSet.Close

      AutomaticallyCreateHeatsAndFinals = True
    End If

  End If

AutomaticallyCreateHeatsAndFinals_Exit:
  
  ReturnValue = SysCmd(SYSCMD_REMOVEMETER)
  DoCmd.RunMacro "ClosePleaseWait"
  Exit Function
  
AutomaticallyCreateHeatsAndFinals_Err:
  AutomaticallyCreateHeatsAndFinals = False
  MsgBox "An error occurred in [AutomaticallyCreateHeatsAndFinals]: " & Err.Description, vbCritical
  Resume AutomaticallyCreateHeatsAndFinals_Exit
  
End Function

' **********************************************************************************************************************************************************
' * Routine creates entries in the Lane Template table which is then used to generate reports
' * It is run after the lanes value is modified in the EventType details form
' **********************************************************************************************************************************************************

Public Sub UpdateLaneTemplate(ET_Code As Long, Lane_Cnt As Variant)

  Dim Q As String, i As Integer, db As Database, rs As Recordset
    
  Q = "DELETE DISTINCTROW [Lane Template].ET_Code FROM [Lane Template]"
  Q = Q & " WHERE [Lane Template].ET_Code=" & ET_Code

  DoCmd.SetWarnings False
  DoCmd.RunSQL Q
  DoCmd.SetWarnings True

  Set db = CurrentDb
  Set rs = db.OpenRecordset("Lane Template", DB_OPEN_DYNASET)   ' Create Recordset.

  For i = 1 To Lane_Cnt

      rs.AddNew
      rs!ET_Code = ET_Code
      rs![Lanes] = i
      rs.Update
               
  Next i

  rs.Close

End Sub