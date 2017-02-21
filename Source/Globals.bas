Option Compare Database   'Use database order for string comparisons
Option Explicit

Global DoUpdateEventCompetitorAge As Boolean

Global In_Val As Variant
Global TotRecs As Long
Global LastTime As Variant


Global Const Debugging = False

Global Const HaveGraphs = True

Global Delm As String
Global Const LightBlue = 16777088
Global Const White = 16777215
Global Const LightRed = 8421631
Global Const DarkGrey = 8421504
Global CalculatePlaces As Variant
Global SexFormat As Variant
Global HeatFormat As Variant


Global ChosenFinalHouse As Variant

Global OpenFormType As Variant

Global ShowDialog As Variant
Global GlobalCancel As Variant
Global GlobalNo As Variant
Global GlobalVariable As Variant
Global GlobalChange As Variant
Global DontEditPromotionFinalsMessage As Variant
Global PleaseWaitMsg As Variant
Global ReturnVar As Variant
Global Response As Variant
Global UserQuit As Boolean
Global GlobalPlaceChange As Boolean

Global Q As String

Global DisplayRecords As Integer

Global Const DEMO = False
Global Const DEMOcompetitors = 75
Global Const DEMOmessage = "This is a demonstration version of the Sports Administrator.  The demo version is limited to 75 competitors.  If you wish to purchase the full version of Sports Administrator, please read the About information from the Main Menu."
Global Const DEMOmessage2 = "This is a demonstration version of the Sports Administrator.  The carnival you are attempting to make active has more than 75 competitors.  The demo version is limited to 75 competitors.  If you wish to purchase the full version of Sports Administrator, please read the About information from the Main Menu."
Global Const aDebug = False
Global EventAgeArray() As String

Global AlwaysOpenRS As Recordset

Type HouseComp
    H As String
    C As String
    Hid As Long
End Type

'*****************************************************************************************************************************
'Purpose:       -
'Parameters:    None
'Returns:       None
'Created By:    Andrew Rogers
'Created On:    Sun 16/Feb/2003
'Comments:      None
'*****************************************************************************************************************************
Public Function AgeFilter(HeatAge)
On Error GoTo AgeFilter_Err

    Dim Length As Variant

    If UCase(Right(HeatAge, 2)) = "_U" Then
        
        Length = Len(HeatAge)
        AgeFilter = "<=" & Val(Left(HeatAge, Length - 2))

    ElseIf UCase(Right(HeatAge, 2)) = "_O" Then
        
        Length = Len(HeatAge)
        AgeFilter = ">=" & Val(Left(HeatAge, Length - 2))

    ElseIf HeatAge = "OPEN" Then
        AgeFilter = " Like """ & "*"""

    Else
        AgeFilter = "=" & Val(HeatAge)
    End If
    



AgeFilter_Exit:
  On Error Resume Next
  Exit Function

AgeFilter_Err:
  Call DisplayErrMsg("AgeFilter")
  Resume AgeFilter_Exit

End Function

Function Better(Res1, Ecode)

On Error GoTo Better_Err
    Dim U As Variant, Order As Variant, ET_Code As Long

    ' Determines whether a given result is better than
    
    
    If Not IsNull(DLookup("[nResult]", "Records", "[E_Code]=" & Ecode)) Then
        ET_Code = DLookup("[ET_Code]", "Events", "[E_Code]=" & Ecode)
        U = DLookup("[Units]", "EventType", "[ET_Code]=" & ET_Code)
        Order = DLookup("[Order]", "Units", "[DisplayUnit]=""" & U & """")
        
        Better = False
        If Order = "ASC" Then
            If Res1 <= DMin("[nResult]", "Records", "[E_Code]=" & Ecode) Then
                Better = True
            End If
        Else
            If Res1 >= DMax("[nResult]", "Records", "[E_Code]=" & Ecode) Then
                Better = True
            End If
    
        End If
    Else
        Better = True
    End If
         
Better_Exit:
  Exit Function
  
Better_Err:
  MsgBox ("An error has occured in [Better]: " & Err.Description)
  GoTo Better_Exit
  
End Function

Function CalcResult(Unit As String, Power As Integer, Valu As String, Delm As String, nValu As String, _
                    i As Integer, AddZero As Integer, ByRef success As Boolean) As Double

    Dim Mult As Variant
    
    success = True
    
    Select Case Unit

        Case "SECS" ' Seconds
            
            Select Case Power

                Case 1
                    CalcResult = Val(Valu)
                    Delm = "."
                    nValu = Left$(nValu, i - 1) & ".0"
                Case 2
                    CalcResult = Val("." & Valu)
                    Delm = "?"

                Case Else
                    MsgBox "The value you entered is not recognised.  Please check that you have entered the value correctly.", 32
                    Delm = "?"
                    
            End Select 'Power

        Case "MINS" ' Minutes

            Select Case Power

                Case 1 ' Mins
                    CalcResult = Val(Valu) * 60
                    Delm = "'"
                    nValu = Left$(nValu, i - 1) & "'0.0"
                Case 2 ' Secs
                    If Val(Valu) > 60 Then
                      MsgBox "That seconds part cannot be greater than 60.", vbInformation
                      success = False
                    End If
                    CalcResult = Val(Valu)
                    Delm = "."
                    nValu = Left$(nValu, i - 1) & ".0"
                Case 3  ' Hsecs
                    CalcResult = Val("." & Valu)
                    Delm = "?"
                    
                Case Else
                    MsgBox "The value you entered is not recognised.  Please check that you have entered the value correctly.", 32
                    Delm = "?"

            End Select 'Power

        Case "HRS" 'Hours

            Select Case Power

                Case 1
                    CalcResult = Val(Valu) * 60 * 60
                    Delm = """"
                    nValu = Left$(nValu, i - 1) & """00'00.00"
                Case 2
                    CalcResult = Val(Valu) * 60
                    Delm = "'"
                    nValu = Left$(nValu, i - 1) & "'00.00"
                
                Case 3
                    CalcResult = Val(Valu)
                    Delm = "."
                    nValu = Left$(nValu, i - 1) & ".00"
                Case 4
                    CalcResult = Val("." & Valu)
                    Delm = "?"
                    'AddZero = 2 - Len(Valu)
                    
                Case Else
                    MsgBox "The value you entered is not recognised.  Please check that you have entered the value correctly.", 32
                    Delm = "?"

            End Select 'Power

        Case "M" 'Meters

            Select Case Power

                Case 1
                    CalcResult = Val(Valu)
                    Delm = "."
                    nValu = Left$(nValu, i - 1) & ".0"
                Case 2
                    CalcResult = Val("." & Valu)
                    Delm = "?"
                    
                Case Else
                    MsgBox "The value you entered is not recognised.  Please check that you have entered the value correctly.", 32
                    Delm = "?"

            End Select 'Power

        Case "KM" 'Meters

            Select Case Power

                Case 1
                    CalcResult = Val(Valu) * 1000
                    Delm = "."
                    nValu = Left$(nValu, i - 1) & ".0"
                Case 2
                    Mult = 10 ^ (3 - Len(Trim(Valu)))
                    CalcResult = Val(Valu) * Mult
                    Delm = "?"
                    
                Case Else
                    MsgBox "The value you entered is not recognised.  Please check that you have entered the value correctly.", 32
                    Delm = "?"

            End Select 'Power

        Case "PTS" 'Points

            Select Case Power

                Case 1
                    CalcResult = Val(Valu)
                    Delm = "."
                    nValu = Left$(nValu, i - 1) & ".0"
                Case 2
                    CalcResult = Val("." & Valu)
                    Delm = "?"

                Case Else
                    MsgBox "The value you entered is not recognised.  Please check that you have entered the value correctly.", 32
                    Delm = "?"

            End Select 'Power

        Case Else
            MsgBox ("The unit is not recognised.  Reselect the unit in the Event Details form for the event that you are currently working on.")

    End Select ' Unit

End Function

Public Function Calculate_Competitor_Lane(E_Code, F_Lev, H_Code, Heat)

    On Error GoTo CC_Err

    Dim LaneAllocated As Variant, AllocatedLane As Variant, Q As Variant

    ' Determine if it is the lowest final level
    ' If So then
    '
    ' Find first lane allocated to that house
    '   Check if it is used
    '   IF SO then
    '

    '   Else Assign this lane
    '

    Dim Criteria As String, db As Database, rs As Recordset, LRS As Recordset
    
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset("SELECT * FROM Heats ORDER BY [F_Lev] Desc", DB_OPEN_DYNASET)   ' Create Recordset.
    Set LRS = db.OpenRecordset("Lanes", DB_OPEN_DYNASET)   ' Create Recordset.
    
    Criteria = "[E_Code]=" & E_Code
    rs.FindFirst Criteria

    AllocatedLane = 0
    
    If rs!F_Lev = F_Lev Then 'Lowest Final Level
            
        Criteria = "[H_Code] = " & H_Code
        LRS.FindFirst Criteria

        LaneAllocated = False
        While Not LRS.EOF And Not (LRS.NoMatch) And Not (LaneAllocated)
            AllocatedLane = LRS!Lane
            Q = "[E_Code]=" & E_Code & " AND [F_Lev]=" & F_Lev & " AND [Heat]=" & Heat & " AND [Lane]=" & AllocatedLane
            If IsNull(DLookup("[Lane]", "CompEvents", Q)) Then ' Lane is available
                LaneAllocated = True
            Else
                LRS.FindNext Criteria
                AllocatedLane = 0
            End If


        Wend
    
        Calculate_Competitor_Lane = AllocatedLane

    Else
        If DontEditPromotionFinalsMessage Then
            Response = MsgBox("This final-level is not the lowest final-level.  The competitors and lanes for this final-level should be determined when the competitors are promoted.", vbOKOnly + vbInformation)
            DontEditPromotionFinalsMessage = False
        End If

    End If

    rs.Close
    LRS.Close
    

CC_Exit:
    Exit Function

CC_Err:
    MsgBox Error$
    GoTo CC_Exit
    
End Function

Sub Calculate_Results(res As String, nValu As String, Runit As String, ByRef success As Boolean)

    ' ---wrong I think ---------------------------------------------------------------------------------------
    ' - Res = is the text value enterd by the user representing the result gained by the competitor
    ' -     = It is returned as a String value in the correct unit format
    ' - nValu = Comes in as NULL and is returned as the Numeric Result
    ' ------------------------------------------------------------------------------------------

    ' ------------------------------------------------------------------------------------------
    ' - Res = is the text value enterd by the user representing the result gained by the competitor
    ' -     = It is returned as as the Numeric Result
    ' - nValu = Comes in as NULL and is returned in the correct unit format
    ' ------------------------------------------------------------------------------------------
    'Dim Runit As String
    Dim cUnit As String
    Dim Valu As String
    Dim nRes As Double
    Dim Power As Integer
    Dim Delm As String
    Dim i As Integer
    Dim AddZero As Integer
    Dim Order As Variant
    Dim ResLength As Variant
    Dim Char As Variant
    Dim FirstNN As Variant
    Dim SecNN As Variant
    Dim LeftRes As Variant
    Dim RightRes As Variant
    
  If UCase(Left(res, 1)) = "F" Then
    nValu = "FOUL"
    Order = DLookup("[Order]", "Units", "[DisplayUnit]=""" & Runit & """")
    If Order = "Asc" Then
        nRes = 3E+38
    Else
        nRes = -1E+38
    End If

  ElseIf UCase(Left(res, 1)) = "P" Then
    nValu = "PARTICIPATE"
    Order = DLookup("[Order]", "Units", "[DisplayUnit]=""" & Runit & """")
    If Order = "Asc" Then
        nRes = 3E+38
    Else
        nRes = -1E+38
    End If

  Else

    'Runit = Forms![EnterCompetitors]![EC_Subform].Form![Unit]
    'Res = Forms![EnterCompetitors]![EC_Subform].Form![Res]
    
    res = Trim(res)
    ResLength = Len(res)
    Runit = Trim(Runit)

    nRes = 0   ' New Result is in SECONDS

    cUnit = UCase(Runit)

    Power = 0
    Valu = ""
    nValu = ""
    AddZero = 0
    i = 0

    For i = 1 To ResLength Step 1
        
      Char = Mid$(res, i, 1)

      ' **** The Character is a delimter of some sort ****
      If Not (IsNumeric(Char)) Or Char = "." Then
          Power = Power + 1
          nValu = res
          nRes = nRes + CalcResult(cUnit, Power, Valu, Delm, nValu, i, AddZero, success)
          Valu = ""
          If Not success Then GoTo ResultFormatError
          Mid$(res, i, 1) = Delm

      Else
          Valu = Valu + Char
      End If
      
    Next i

    Power = Power + 1
    nValu = res
    nRes = nRes + CalcResult(cUnit, Power, Valu, Delm, nValu, i, AddZero, success)
    If Not success Then GoTo ResultFormatError
    
    i = 0
    FirstNN = 0
    SecNN = 0

    While i < Len(nValu)
        i = i + 1
        If Not (IsNumeric(Mid$(nValu, i, 1))) Or Mid$(nValu, i, 1) = "." Then
            FirstNN = SecNN
            SecNN = i
        End If

        While (FirstNN <> 0) And (FirstNN + 3 > SecNN)
            LeftRes = Left$(nValu, FirstNN)
            RightRes = Right$(nValu, Len(nValu) - FirstNN)
            nValu = LeftRes & "0" & RightRes '
            i = i + 1
            SecNN = SecNN + 1
        Wend
        
    Wend

    While SecNN + 2 > Len(nValu)
        nValu = nValu & "0"
    Wend

  End If
    
  res = Str(nRes)
  success = True
  
Calculate_Results_Exit:
  Exit Sub
  
ResultFormatError:
  success = False
  GoTo Calculate_Results_Exit

End Sub

Function CalculatePercTotal(T, H, P)

    If P > 0 Then
        CalculatePercTotal = Format(T / P * 100, "0.0") & " (" & P & ")"
    Else
        CalculatePercTotal = 0
    End If

End Function

Function CarnivalDir(RD)
    
    ' Determines the full Carnival Directory from the 'Relative' directory stored in the Carnivals Table.

    If IsNull(RD) Then
        ' Carnival file is originally located in the sports.mdb directory
        CarnivalDir = DBPath()
    ElseIf Mid$(RD, 2, 2) = ":\" Then
        ' Carnival file is specified by an absoulute directory
        CarnivalDir = RD
    ElseIf Left$(RD, 2) = "\\" Then
        ' Carnival file is specified by a UNC path
        CarnivalDir = RD
    Else
        ' Carnival file is relative to the sports.mdb directory (say 'carnival\')
        CarnivalDir = DBPath() & RD
    End If
    
    
End Function

Function CheckFinalIntegrity(code, T)

      
    Dim LargestFinal As Variant, F As Variant
    CheckFinalIntegrity = True
    If Not IsNull(code) Then
    
        If T = "HEATS" Then
            LargestFinal = DMax("[F_Lev]", "Heats", "[E_Code]=" & code)
        Else
            LargestFinal = DMax("[F_Lev]", "Final_Lev", "[ET_Code]=" & code)
        End If
        If Not IsNull(LargestFinal) Then
          For F = 0 To LargestFinal
            If T = "HEATS" Then
                If IsNull(DLookup("[F_Lev]", "Heats", "[E_Code]=" & code & " AND [F_Lev]=" & F)) Then
                    CheckFinalIntegrity = False
                    GoTo CheckFinalIntegrityExit
                End If
            Else
                If IsNull(DLookup("[F_Lev]", "Final_Lev", "[ET_Code]=" & code & " AND [F_Lev]=" & F)) Then
                    CheckFinalIntegrity = False
                    GoTo CheckFinalIntegrityExit
                End If
            End If
          Next F
        End If
    End If

CheckFinalIntegrityExit:
    Exit Function

End Function

Sub CheckIfRecordBroken(E_Code, Heat, F_Lev)
On Error GoTo Err_CheckIfRecordBroken
    'Stop

    Dim U As Variant, Order As Variant, Res1 As Variant, Better   As Variant, Criteria As Variant
    Dim rs As Recordset, db As Database, Criteria2 As Variant, Q As Variant, ValuesText As Variant
    Dim Fullname As Variant, Response As Variant, AlertToRecord As Variant, Result As Variant

    'AlertToRecord = DLookup("[AlertToRecord]", "MiscellaneousLocal") 'not used presently
    
    U = DLookup("[Units]", "Events in Full", "[E_Code]=" & E_Code)
    Order = DLookup("[Order]", "Units", "[DisplayUnit]=""" & U & """")
    
    If Heat = -1 And F_Lev = -1 Then
        ' This is not possible under normal circumstances
        ' Set this so that the routine checks an entire event at a time rather than just a heat.
        ' However for individual races check only the heat that has been entered.
        ' Checking just a heat is probably overkill but it is more logical to the person entering the data

        Criteria = "[E_Code]=" & E_Code & " AND [EffectsRecords]=TRUE "
    Else
        Criteria = "[E_Code]=" & E_Code & " AND [Heat]=" & Heat & " AND [F_Lev]=" & F_Lev & " AND [EffectsRecords]=TRUE "
    End If

    If Order = "ASC" Then
        Res1 = DMin("[nResult]", "CompEvents-Records", Criteria & "AND nResult<>0")
        
    Else
        Res1 = DMax("[nResult]", "CompEvents-Records", Criteria)

    End If
    
    ' Ensure that there is a result for the event
    If Not (IsNull(Res1) Or Res1 <= 0) Then
    
        If Not IsNull(DLookup("[nResult]", "Records", "[E_Code]=" & E_Code)) Then
            ' There is a previous record
            If Order = "ASC" Then
                If Res1 <= DMin("[nResult]", "Records", "[E_Code]=" & E_Code) Then
                    Set rs = CurrentDb.OpenRecordset("CompEvents-with Competitor Names", DB_OPEN_DYNASET)   ' Create dynaset.
                    Criteria = Criteria & " AND [nResult] = " & Res1
                    rs.FindFirst Criteria
                    While Not (rs.EOF Or rs.NoMatch)
                        GoSub AddCompetitorToRecords
                        rs.FindNext Criteria
                    Wend
    
                End If
    
            Else
                If Res1 >= DMax("[nResult]", "Records", "[E_Code]=" & E_Code) Then
                    Set rs = CurrentDb.OpenRecordset("CompEvents-with Competitor Names", DB_OPEN_DYNASET)   ' Create dynaset.
                    Criteria = Criteria & " AND [nResult] = " & Res1
                    rs.FindFirst Criteria
                    While Not (rs.EOF Or rs.NoMatch)
                        GoSub AddCompetitorToRecords
                        rs.FindNext Criteria
                    Wend
                    
                End If
        
            End If
        Else
            ' There has been no previous record set
    
            Set db = DBEngine.Workspaces(0).Databases(0)
            Set rs = db.OpenRecordset("CompEvents-with Competitor Names", DB_OPEN_DYNASET)   ' Create dynaset.
            Criteria = Criteria & " AND [nResult] = " & Res1
            rs.FindFirst Criteria
            While Not (rs.EOF Or rs.NoMatch)
                GoSub AddCompetitorToRecords
                rs.FindNext Criteria
            Wend
        
        End If
    End If
    GoTo Exit_CheckIfRecordBroken

'**********************************************************
AddCompetitorToRecords:

  Q = rs!Gname & " " & rs!Surname & " has set a new record for this event (" & rs!Result & " " & U & ").  " & LFCR & LFCR
  Q = Q & "Do you wish to accept it?"
  
  Response = MsgBox(Q, vbYesNo + vbDefaultButton2 + vbQuestion, "New Record")
  
  If Response = vbYes Then
    'Criteria2 = "[E_Code]=" & E_Code & " AND [Surname]= """ & RS![Surname] & """ AND [Gname]=""" & RS!Gname & """ AND [H_Code]=""" & RS![H_Code] & """ AND [Date]= #" & Format$(Now, "mm/dd/yyyy") & "# AND [nResult] = " & RS!nResult
    Criteria2 = "[E_Code]=" & E_Code & " AND [Surname]= """ & rs![Surname] & """ AND [Gname]=""" & rs!Gname & """ AND [H_Code]=""" & rs![H_Code] & """ AND [nResult] = " & rs!nResult
    
    If IsNull(DLookup("[E_Code]", "Records", Criteria2)) Then
        ' Competitor has not already been added
        
        If IsNull(rs!Result) Then
            Result = 0
        Else
            Result = rs!Result
        End If

        ValuesText = "(" & E_Code & ",""" & rs!Surname & """,""" & rs!Gname & """,""" & rs!H_Code & """, #" & Format$(Now, "mm/dd/yyyy") & "# ," & rs!nResult & ",""" & Result & """)"
        Q = "INSERT INTO Records ( E_Code, Surname, Gname, H_Code, [Date], nResult, Result ) "
        Q = Q & "VALUES " & ValuesText
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True

        If GlobalVariable Then
            Forms!EnterCompetitors!Record = Result
            Forms!EnterCompetitors!nRecord = rs!nResult
        Else
            Q = "UPDATE DISTINCTROW Events SET Events.Record = """ & Result & """, Events.nRecord = " & rs!nResult
            Q = Q & " WHERE Events.E_Code=" & E_Code
            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
        End If
        
        'If GlobalVariable And AlertToRecord Then
        '    Fullname = rs!Gname & " " & UCase(rs!Surname)
        '    Response = MsgBox(Fullname & " has set a new record for this event.", 0, "New Record")
        'End If
      End If
    End If

    Return


Exit_CheckIfRecordBroken:
    Exit Sub
    
Err_CheckIfRecordBroken:
    MsgBox ("Error in CheckiIfRecordBroken:" & Error$)
    GoTo Exit_CheckIfRecordBroken


End Sub

Function CompetitorsInEvent(Ecode, FLev, Heat)

    CompetitorsInEvent = DCount("[PIN]", "CompEvents", "[E_Code]=" & Ecode & " AND [F_Lev]=" & FLev & " AND [Heat]=" & Heat)

End Function

Public Function DetermineAge(Eage As String) As Byte
On Error GoTo DetermineAge_Err
'Converts an EventAge into a numerical age

    Dim TempAge As Variant, CurYear As Variant

    If IsDate(Eage) Then
        'DetermineAge = eAge
        DetermineAge = Year(Now) - Year(Eage)
            
    ElseIf Eage = "OPEN" Then
                  
        CurYear = Year(Now)
        'tComp!DOB = "1/1/11"
        DetermineAge = DLookup("[OpenAge]", "Miscellaneous")
            
    Else
      DetermineAge = Val(Nz(Eage))
            
    End If



DetermineAge_Exit:
  On Error Resume Next
  Exit Function

DetermineAge_Err:
  Call DisplayErrMsg("DetermineAge")
  Resume DetermineAge_Exit

End Function

Function OLDDetermineAge(Eage As String)
'Converts an EventAge into a numerical age

    Dim TempAge As Variant, CurYear As Variant

    If IsDate(Eage) Then
        'DetermineAge = eAge
        TempAge = Year(Now) - Year(Eage)
        If TempAge >= DLookup("[OpenAge]", "Miscellaneous") Then
            OLDDetermineAge = "OPEN"
        Else
            OLDDetermineAge = Trim(Str(TempAge))
        End If
            
    ElseIf Eage = "OPEN" Then
                  
        CurYear = Year(Now)
        'tComp!DOB = "1/1/11"
        OLDDetermineAge = "OPEN"
            
    Else
        'tComp!DOB = "1/1/" & Year(Now) - eAge
        If Not (IsNull(Eage)) Then
            If Val(Eage) >= DLookup("[OpenAge]", "Miscellaneous") Then
                OLDDetermineAge = "OPEN"
            
            Else
                OLDDetermineAge = Eage
            End If

        End If
            
    End If


End Function

Function DetermineAge_ImportCompetitors(DOB As Variant, CutDay As Integer, CutMonth As Integer)

  ' Should have already trapped for DOB being null
  Dim Cage As String
  
  If Not IsDate(DOB) Then
    'MsgBox ("The Date of Birth is not a valid date.")
    Cage = ""
  Else
    
    Dim TempAge As Variant, CurYear As Variant
    Dim Cday As Integer, Cmonth As Integer, Cyear As Integer
    
    Cday = Format(DOB, "dd") ' Day competitor was born
    Cmonth = Int(Format(DOB, "mm")) ' Month competitor was born
    Cyear = Int(Format(DOB, "yyyy")) ' Year competitor was born
    
    CurYear = Int(Format(Now, "yyyy")) ' CurYear
    
    If Cmonth > CutMonth Then
      Cage = Str(CurYear - Cyear)
    ElseIf Cmonth < CutMonth Then
      Cage = Str(CurYear - Cyear) + 1
    Else ' Born in the same month as the CutOff month
      If Cday >= CutDay Then
        Cage = Str(CurYear - Cyear)
      Else
        Cage = Str(CurYear - Cyear) + 1
      End If
    End If
  End If
  
  DetermineAge_ImportCompetitors = Cage

End Function

Function DetermineDOB(Eage)

    Dim CurYear As Variant
    
    If IsNull(Eage) Then
        DetermineDOB = Null
    ElseIf IsDate(Eage) Then
        DetermineDOB = Eage
            
    ElseIf Eage = "OPEN" Then
                  
        CurYear = Year(Now)
        DetermineDOB = "1/1/1901"
        
    Else
        DetermineDOB = "1/1/" & Year(Now) - Val(Eage)
    End If



End Function

Function DetermineEventAge(A)
On Error GoTo DetermineEventAge_Err

    ' Determines what Event age bracket a competitros age falls into. ie 8 year old in the 09_U age

    Dim Criteria As String, db As Database, rs As Recordset, Q As Variant, Continue  As Variant
    Dim i As Variant, Eage  As Variant, AgeFil As Variant, AQ As Variant
    
    Q = "SELECT DISTINCT Events.Age FROM Events"
    Set rs = CurrentDb.OpenRecordset(Q, DB_OPEN_DYNASET)   ' Create dynaset.
    
    'Stop

    Continue = True
    DetermineEventAge = "UNKNOWN"
    i = 0
    
    'Do Until EventAgeArray(i) = "THATSIT" Or Not Continue

    Do Until (rs.EOF Or Not Continue)    ' Loop until no matching records.
        Eage = rs!Age
        ' Check that there is no age group that this age excludes.  For exaple
        ' 12_O excludes the age bracket 13 or 13_O.  9_U excludes 8_U.  We want to use only the
        ' outermost ones.
        If Outermost(Eage) Then
            AgeFil = AgeFilter(Eage)
            AQ = "(Events.Age=""" & Eage & """) AND (" & Val(A) & AgeFil & ")"
        
            If DCount("[Age]", "EventAges", AQ) > 0 Then
               DetermineEventAge = Eage
               Continue = False
            End If
        End If
        rs.MoveNext
        'i = i + 1
    Loop

    'If DetermineEventAge = "UNKNOWN" Then DetermineEventAge = A
    
DetermineEventAge_exit:
    Exit Function

DetermineEventAge_Err:
    MsgBox ("DetermineEventAge:" & Error$)
    GoTo DetermineEventAge_exit
    
End Function

Function DetermineFullName(s, g)

    If IsNull(s) And IsNull(g) Then
        DetermineFullName = ""
    Else
        DetermineFullName = UCase(s) & ", " & g
    End If

End Function

Function DetermineH_ID(H_Code)

    DetermineH_ID = DLookup("[H_ID]", "House", "[H_Code]=""" & H_Code & """")

End Function

Function DetermineHeat(Heat)

    If IsNumeric(Heat) Then
        DetermineHeat = Heat
    Else
        DetermineHeat = Asc(UCase(Heat)) - 64
    End If

End Function

Function DetermineLane(E_Code, Place)

    Dim ET_Code As Variant, Lane  As Variant

    ET_Code = EventTypeID(E_Code)
    Lane = DLookup("[Lane]", "Lane Promotion Allocation", "[ET_Code]=" & ET_Code & " AND [Place]=" & Place)

    If IsNull(Lane) Then
        DetermineLane = 0
    Else
        DetermineLane = Lane
    End If
         
End Function

Function DeterminePoints(PL, PtScale)

    If IsNull(PL) Then
        DeterminePoints = 0
    ElseIf IsNull(PtScale) Then
        MsgBox ("Error: unassigned PointScale")
    Else
        DeterminePoints = DLookup("[Points]", "PointsScale", "[Place]=" & PL & " AND [PtScale]=""" & PtScale & """")
    End If
    
End Function

Function DetermineSex(Sex)
    
    Dim Fsex As Variant

    Fsex = UCase(Left(Sex, 1))
    If Fsex = "B" Or Fsex = "M" Then
        DetermineSex = "M"
    ElseIf Fsex = "G" Or Fsex = "F" Then ' G or F
        DetermineSex = "F"
    Else
        DetermineSex = False
    End If
        

End Function

Function DisplayPoints(Pt)

    If Pt = 0 Then
        DisplayPoints = Null
    Else
        DisplayPoints = Pt
    End If

End Function

Function DisplayRecHolder(N, H)

    If IsNull(H) Then H = -1
    If IsNull(N) Then
        DisplayRecHolder = "Record Holder: " & DLookup("[H_Name]", "House", "[H_ID]=" & H)
    Else
        DisplayRecHolder = "Record Holder: " & Trim(N) & " / " & DLookup("[H_Name]", "House", "[H_ID]=" & H)
    End If

End Function

Function DisplayResult(res)
    
    If IsNull(res) Then
        DisplayResult = ""
    ElseIf res = "0" Then
        DisplayResult = ""
    Else
        DisplayResult = res
    End If
        
    
End Function

Function EventAge(E_Code)

    EventAge = DLookup("[Age]", "Events", "[E_Code]= " & E_Code)

End Function

Function EventDescription(E_Code)

    Dim ET_Code As Variant

    ET_Code = DLookup("[ET_Code]", "Events", "[E_Code]= " & E_Code)
    EventDescription = DLookup("[ET_Des]", "EventType", "[ET_Code]= " & ET_Code)
    
End Function

Function EventSex(E_Code)

    EventSex = DLookup("[Sex]", "Events", "[E_Code]= " & E_Code)
    
End Function

Function EventTypeID(E_Code)

    EventTypeID = DLookup("[ET_Code]", "Events", "[E_Code] = " & E_Code)

End Function

Function FinalHouse()

    FinalHouse = ChosenFinalHouse
    'FinalHouse = "Muel"

End Function

Function FindLastEntry(uTable, uField As Field)

    Dim db As Database, rs As Recordset
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset(uTable, DB_OPEN_DYNASET)   ' Create Recordset.
    
    rs.MoveLast
    FindLastEntry = rs!uField
    
    rs.Close

End Function

Function FormatGname(N)

    Dim L As Variant, FirstLetter As Variant

    L = Len(N)
    FirstLetter = Left$(N, 1)
    FirstLetter = UCase$(FirstLetter)
    N = FirstLetter & LCase$(Mid$(N, 2, L - 1))
    FormatGname = N

End Function

Function GenerateAgeFilterOLD(A)
    'Stop

    Dim Age As Variant, Q As Variant

    Age = AgeFilter([Forms]![EnterCompetitors]![AgeFld])
    
    Q = "SELECT UCase(Trim([Surname]))+""" & ", " & """+Trim([Gname]) AS fName, CompetitorsOrdered.H_Code, CompetitorsOrdered.PIN, House.Include "
    Q = Q & "FROM House INNER JOIN CompetitorsOrdered ON House.H_Code = CompetitorsOrdered.H_Code "
    Q = Q & "WHERE ((CompetitorsOrdered.Sex= """ & [Forms]![EnterCompetitors]![SexFld] & """)) AND (val(CompetitorsOrdered.Age)" & Age & " ) AND (House.Include=Yes) " 'AND CompetitorsOrdered.Flag = True ORDER by [Order]"
    Q = Q & "ORDER BY [Surname], [Gname] "
    
    'GenerateAgeFilter = Q

End Function

Function GenerateAgeFilter(Age, Sex)
    'Stop

    Dim Q As Variant
    Age = AgeFilter(Age)
    
    Q = "SELECT UCase(Trim([Surname]))+""" & ", " & """+Trim([Gname]) AS Name, CompetitorsOrdered.H_Code as Team, Sex, CompetitorsOrdered.PIN, [Age], House.Include "
    Q = Q & "FROM House INNER JOIN CompetitorsOrdered ON House.H_Code = CompetitorsOrdered.H_Code WHERE "
    If Sex <> "-" Then
      Q = Q & " (CompetitorsOrdered.Sex= """ & Sex & """) AND "
    End If
    Q = Q & " (val(CompetitorsOrdered.Age)" & Age & " ) AND (House.Include=Yes) " 'AND CompetitorsOrdered.Flag = True ORDER by [Order]"
    Q = Q & "ORDER BY [Surname], [Gname] "
    
    GenerateAgeFilter = Q

End Function

Function GenerateSexFilterOLD()

    Dim Q As Variant

    Q = "SELECT UCase(Trim([Surname]))+""" & ", " & """+Trim([Gname]) AS Name, CompetitorsOrdered.H_Code as Team, Sex, CompetitorsOrdered.PIN "
    Q = Q & "FROM CompetitorsOrdered "
    Q = Q & "WHERE ((CompetitorsOrdered.Sex= """ & [Forms]![EnterCompetitors]![SexFld] & """))  " 'AND CompetitorsOrdered.Flag = True ORDER by [Order]"
    Q = Q & "ORDER BY [Surname], [Gname] "
    
    'GenerateSexFilter = Q

End Function

Function GenerateSexFilter(Sex)

    Dim Q As Variant

    Q = "SELECT UCase(Trim([Surname]))+""" & ", " & """+Trim([Gname]) AS Name, CompetitorsOrdered.H_Code as Team, Sex, CompetitorsOrdered.PIN, [Age] "
    Q = Q & "FROM CompetitorsOrdered "
    If Sex <> "-" Then
      Q = Q & " WHERE CompetitorsOrdered.Sex= """ & Sex & """ " 'AND CompetitorsOrdered.Flag = True ORDER by [Order]"
    End If
    Q = Q & " ORDER BY [Surname], [Gname] "
    
    GenerateSexFilter = Q

End Function

Function GetCarnivalFile(C)
    
    Dim CD As Variant

    CD = ExtractDirectory(C)
    GetCarnivalFile = Right$(C, Len(C) - Len(CD))

End Function

Function GetCarnivalFullDir(C)
            
    Dim FD As Variant

    FD = ExtractDirectory(C)
    If Not IsNull(FD) Then
        If Mid$(FD, 2, 2) = ":\" Then
            GetCarnivalFullDir = FD
        ElseIf Left$(FD, 2) = "\\" Then
            ' UNC path used
            GetCarnivalFullDir = FD
        Else
            GetCarnivalFullDir = DBPath() & FD
        End If

    Else
        GetCarnivalFullDir = DBPath() & FD
    End If
    
End Function


Public Function SportAddErrorCode()
On Error GoTo SportAddErrorCode_Err

  If CurrentUser = "Owner" Then Call AddErrorCheckingCode
  
SportAddErrorCode_Exit:
  On Error Resume Next
  Exit Function

SportAddErrorCode_Err:
  Call DisplayErrMsg("SportAddErrorCode")
  Resume SportAddErrorCode_Exit

End Function

Function GetCarnivalRelDir(FullCF)

    
    Dim DBp As Variant, NewDir As Variant

    DBp = DBPath()
    If DBp = Left$(ExtractDirectory(FullCF), Len(DBp)) Then
        NewDir = ExtractDirectory(Right$(FullCF, Len(FullCF) - Len(DBp)))
    Else
        NewDir = ExtractDirectory(FullCF)
    End If
    If IsNull(NewDir) Then
        GetCarnivalRelDir = ""
    Else
        GetCarnivalRelDir = NewDir
    End If

    
End Function

Function InitialiseWaitMessage()

'    If Not IsNull(SysCmd(acSysCmdProfile)) Then
'      MsgBox ("Profile= " & SysCmd(acSysCmdProfile))
'    Else
'      MsgBox ("Profile= NONE")
'    End If
    
    PleaseWaitMsg = "Starting the Sports Administrator ..."

End Function

Function oF(N, T)
    
    If T = "M" Then ' Modal Form
        DoCmd.OpenForm N, , , , , A_DIALOG
    Else
        DoCmd.OpenForm N
    End If

End Function

Public Function OpenForm(Fname As String)
On Error GoTo OpenForm_Err
    
    Dim DocName As String
    Dim LinkCriteria As String

    DocName = Fname
    LinkCriteria = ""
    DoCmd.OpenForm DocName, , , LinkCriteria


OpenForm_Exit:
  On Error Resume Next
  Exit Function

OpenForm_Err:
  Call DisplayErrMsg("OpenForm")
  Resume OpenForm_Exit

End Function

Function Outermost(A)

    Dim AageOnly As Variant, AgeCheck As Variant
    'Stop

    ' Checks if an event age is the outermost.  That is 12_O is not the outermost when there exists 13_O
    ' This happens when there is say two events that have a different youngest or oldest age goup.
    ' EG 12_O and 13_O.  There will usually be a 12 age group if there is a 13_O age group.
        
    Outermost = True

    If Right(A, 2) = "_U" Then
        AageOnly = Val(Left(A, Len(A) - 2))
        AgeCheck = DLookup("[Age]", "EventAges", "Val([age]) < " & AageOnly)
        If Not IsNull(AgeCheck) Then
            If Val(AgeCheck) <> 0 Then
                Outermost = False
            End If
        End If
    ElseIf Right(A, 2) = "_O" Then
        AageOnly = Val(Left(A, Len(A) - 2))
        AgeCheck = DLookup("[Age]", "EventAges", "Val([age]) > " & AageOnly)
        If Not IsNull(AgeCheck) Then
            If Val(AgeCheck) <> 0 Then
                Outermost = False
            End If
        End If

    End If


End Function


Function PromoteEventFinal(E_Code)
'On Error GoTo PromoteEventFinal_Err

    ' Determine Final Level to be Promoted (Promote_FL)
    ' Determine Final Level to be Promoted TO (New_FL)
    ' Determine Number of Heats (Num_Heats) in New_FL
    ' Determine Number of Lanes (Num_Lanes) in New_FL
    ' Determine Number of Competitors (Num_Competitors) to promote
    ' Determine promotion type (Pro_Ty)
    ' Determine Time or Place promotion (Time_Pro)
    '
    ' Select all students in that Final Level
    '   Order them appropriately
    '
    'If Pro_Ty = "Smooth" then
    '   Find First Competitor
    '   Find First Heat
    '   For i = 1 to NumCompetitors
    '       For l = 1 to Num_Lanes
    '           Add Competitor to Heat
    '           Find Next Competitor
    '       Next l
    '       Find Next Heat
    '   Next i
    '
    'If Pro_Ty = "Staggered" then
    '   Find First Competitor
    '   Find First Heat
    '   For i = 1 to Num_Lanes
    '       For h = 1 to Num_Heats
    '           Add Competitor to Heat
    '           Find Next Competitor
    '           Find Next Heat
    '       Next h
    '   Next i
    ' Set all old E_Code / Promote_FL pairs to Promoted
'------------------------------------------------------------------
    
    ' Determine Final Level to be Promoted (Promote_FL)
    ' Determine Final Level to be Promoted TO (New_FL)
    ' Determine Number of Heats (Num_Heats) in New_FL
    ' Determine promotion type (Pro_Ty)
    ' Determine Time or Place promotion (Time_Pro)

    PromoteEventFinal = False

    Dim Criteria As String, rs As Recordset, Promote_FL As Variant, Pro_Ty As Variant
    Dim Time_Pro As Variant, Ev As Variant, New_FL As Variant, Num_Heats As Variant
    Dim ET_Code As Variant, LaneCount As Variant, Num_Lanes As Variant, Num_Competitors As Variant
    Dim Q As Variant, uOrder As Variant, Place As Variant, i As Variant, L As Variant, H As Variant
    Dim Units As Variant, Response As Variant, msg As Variant

    Dim EventsRS As Recordset, EventTypeRS As Recordset
    
    'Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Heats ORDER BY [F_Lev] Asc", DB_OPEN_DYNASET)   ' Create Recordset.
    Set EventsRS = CurrentDb.OpenRecordset("Events", dbOpenDynaset)
    Set EventTypeRS = CurrentDb.OpenRecordset("EventType", dbOpenDynaset)
    
    If Not IsNull(DCount("[E_Code]", "Heats", "E_Code = " & E_Code & " AND [Status] = 2")) Then
        Criteria = "E_Code = " & E_Code & " AND [Status] = 2" ' Final Completed
        
        rs.FindFirst Criteria    ' Locate first occurrence.
    
        If rs.NoMatch Then
            MsgBox "There are no finals to promote for this event.", vbExclamation
            PromoteEventFinal = False
        Else
            Promote_FL = rs!F_Lev
            Pro_Ty = rs!Pro_Type
            Time_Pro = rs!UseTimes
            
            Criteria = "E_Code = " & E_Code & " AND [Status] = 1" ' Final Completed
            rs.FindPrevious Criteria
    
            If rs.NoMatch Then  'Beggining of file
                
                'MsgBox ("There are no finals to promote competitors into.  The latest completed final for " & EV & " was the Grand Final.")
                PromoteEventFinal = False
            Else
                New_FL = rs!F_Lev  ' This assumes that the previous F_Lev is the new final level.
                Criteria = "[E_Code] = " & E_Code & " AND [F_Lev] = " & New_FL
                Num_Heats = DCount("[Heat]", "Heats", Criteria)
                
                ' Check if the AllNames flag needs to be set
                If Not IsNull(DLookup("[E_Code]", "Heats", "[E_Code]=" & E_Code & " AND [F_Lev]=" & Promote_FL & " AND [AllNames]=Yes")) Then
                    Q = "UPDATE DISTINCTROW Heats SET Heats.AllNames = Yes WHERE (Heats.E_Code=" & E_Code & " AND Heats.F_Lev=" & New_FL & ")"
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL Q
                    DoCmd.SetWarnings True
                End If
    
    
            ' *** Determine Number of Lanes (Num_Lanes) in New_FL
                
                EventsRS.FindFirst "[E_Code] = " & E_Code
                ET_Code = EventsRS!ET_Code
                'ET_Code = DLookup("[ET_Code]", "Events", "[E_Code] = " & E_Code)
                
                EventTypeRS.FindFirst "[ET_Code] = " & ET_Code
                LaneCount = EventTypeRS![Lane_Cnt]
                'LaneCount = DLookup("[Lane_Cnt]", "EventType", "[ET_Code] = " & ET_Code)
                'Stop
                If LaneCount > 0 Then
                    Num_Lanes = LaneCount
                Else
                    Num_Lanes = DLookup("[ProNum]", "Final_Lev", "[ET_Code]=" & ET_Code & " AND [F_Lev]=" & New_FL)
                    If IsNull(Num_Lanes) Then
                      
                      Ev = DLookup("[ET_Des]", "Events in Full", "[E_Code]=" & E_Code)
                      Ev = Ev & "  Age:" & DLookup("[Age]", "Events in Full", "[E_Code]=" & E_Code)
                      Ev = Ev & "  Sex:" & DLookup("[Sex]", "Events in Full", "[E_Code]=" & E_Code)
            
                      MsgBox ("The number of competitors to be promoted in event " & Ev & " has not been set.  Set this in the SETUP HEATS form.")
                      GoTo Exit_PEF ' PromtionComplete
                    End If
                End If
    
                GoSub DetermineSortOrder
    
                Num_Competitors = Num_Lanes * Num_Heats
        
            ' *** Select all Competitors in that Final Level, Order them appropriately
        
                Dim Crs As Recordset, Hrs As Recordset, NewCRS As Recordset
    
                Q = "SELECT DISTINCTROW CompEvents.*, CompEvents.E_Code, CompEvents.F_Lev "
                Q = Q & "FROM CompEvents "
                Q = Q & "WHERE CompEvents.E_Code= " & E_Code & " AND CompEvents.F_Lev=" & Promote_FL
                Q = Q & " ORDER BY " & uOrder
                              
                Set Crs = CurrentDb.OpenRecordset(Q, DB_OPEN_DYNASET)   ' Create Recordset.
    
                Q = "SELECT DISTINCTROW Heats.E_Code, Heats.F_Lev, Heats.Heat FROM Heats "
                Q = Q & "WHERE Heats.E_Code=" & E_Code & " AND Heats.F_Lev = " & New_FL
                
                Set Hrs = CurrentDb.OpenRecordset(Q, DB_OPEN_DYNASET)   ' Create Recordset.
                Set NewCRS = CurrentDb.OpenRecordset("CompEvents", DB_OPEN_DYNASET)   ' Create Recordset.
    
                Place = 1
                    
                'Stop
    
                If Pro_Ty = "Smooth" Then
                  If Crs.RecordCount = 0 Then
                    'EventDescription
                    
                    Q = "The event cannot be promoted because there are no competitors in the event: "
                    Q = Q & EventDescription(E_Code) & " - "
                    Q = Q & DLookup("[Age]", "Events", "[E_Code]=" & E_Code) & " - "
                    Q = Q & DLookup("[Sex]", "Events", "[E_Code]=" & E_Code) & "."
                    MsgBox Q, vbExclamation
                    
                  Else
                    Crs.MoveFirst
                    Hrs.MoveFirst
                
                    For i = 1 To Num_Heats
                        Place = 1
                        For L = 1 To Num_Lanes
                            GoSub AddCompetitorToHeat
                            Crs.MoveNext
                            If Crs.EOF Then ' There are no more competitors to process
                                GoTo PromtionComplete
                            End If
                                 
                            Place = Place + 1
                        Next L
                        Hrs.MoveNext
                    Next i
                    GoTo PromtionComplete

                  End If
                
                ElseIf Pro_Ty = "Staggered" Then
                  If Crs.RecordCount = 0 Then
                    
                    'MsgBox "No competitors in event"
                    Q = "The event cannot be promoted because there are no competitors in the event: "
                    Q = Q & EventDescription(E_Code) & " - "
                    Q = Q & DLookup("[Age]", "Events", "[E_Code]=" & E_Code) & " - "
                    Q = Q & DLookup("[Sex]", "Events", "[E_Code]=" & E_Code) & "."
                    MsgBox Q, vbExclamation
                  
                  Else
                    Crs.MoveFirst
                                    
                    Place = 1
                    For i = 1 To Num_Lanes
                        Hrs.MoveFirst
                        For H = 1 To Num_Heats
                             GoSub AddCompetitorToHeat
                             Crs.MoveNext
                             If Crs.EOF Then ' There are no more competitors to process
                                 GoTo PromtionComplete
                             End If
    
                             Hrs.MoveNext
                        Next H
                        Place = Place + 1
                    Next i
                    GoTo PromtionComplete
                  End If
                End If
                
                GoTo Exit_PEF ' Should only get to here if there are no competitors in an event that
                              ' was attempted to be promoted
        
PromtionComplete:
                PromoteEventFinal = True
    
                NewCRS.Close
                Crs.Close
                Hrs.Close
    
                GoSub UpdateStatusOfPromotedFinal:
            
            End If   ' *** No finals to promote competitors into
    
        End If ' *** No finals to promote for this event
    End If
    rs.Close

    GoTo Exit_PEF

' ----------------------------------
AddCompetitorToHeat:
    
    NewCRS.AddNew

    NewCRS!PIN = Crs!PIN
    NewCRS!E_Code = E_Code
    NewCRS!Heat = Hrs!Heat
    NewCRS!Lane = 0
    NewCRS!F_Lev = New_FL
    NewCRS!Lane = DetermineLane(E_Code, Place)
    NewCRS.Update
    
    Return
    
' ----------------------------------
DetermineSortOrder:
            
    If Time_Pro Then
      EventTypeRS.FindFirst "[ET_Code] = " & ET_Code
      Units = EventTypeRS!Units
      'Units = DLookup("Units", "EventType", "[ET_Code] = " & ET_Code)
      
      uOrder = DLookup("Order", "Units", "[DisplayUnit] = """ & Units & """")
      uOrder = "[nResult] " & uOrder
    Else
        uOrder = "[Place] ASC"
    End If
            

    Return

' ----------------------------------
UpdateStatusOfPromotedFinal:

  ' Set the status to promoted
  
    Q = "UPDATE DISTINCTROW Heats SET Heats.Status = 3 "
    Q = Q & "WHERE Heats.E_Code= " & E_Code & " AND Heats.F_Lev = " & Promote_FL
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True

    Return


' ----------------------------------
Exit_PEF:
    DoCmd.SetWarnings True
    Exit Function

PromoteEventFinal_Err:
    MsgBox ("Error in PromoteEventFinal: " & Error$)
    GoTo Exit_PEF


End Function


Function PWM()

    'Stop
    PWM = PleaseWaitMsg

End Function

Sub SetCurrentFinal(E_Code)

    On Error GoTo SetCurrentFinal_Error

    DoCmd.SetWarnings True
    Dim Criteria As String, db As Database, rs As Recordset
    Dim LastFinalCompleted As Variant, Cur_Flevel As Variant
    
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset("SELECT * FROM Heats ORDER BY [F_Lev] Desc", DB_OPEN_DYNASET)   ' Create Recordset.
    
    Criteria = "E_Code = " & E_Code & " AND [Completed] = No"
    
    rs.FindFirst Criteria    ' Locate first occurrence.
    
    LastFinalCompleted = False
    
    If DCount("[HE_Code]", "Heats", Criteria) > 0 Then ' Only determine current finals if their are heats already entered
                                                        ' Used to trap no PointsScales potential error
        If rs.NoMatch Then
            rs.MoveLast
            LastFinalCompleted = True
        End If

        Cur_Flevel = rs!F_Lev

        Criteria = "E_Code = " & E_Code
        rs.FindFirst Criteria    ' Locate first occurrence.

        Do Until rs.NoMatch  ' Loop until no matching records.
            rs.Edit          ' Enable editing.
        
            If rs!F_Lev < Cur_Flevel Then
                rs!Status = 0   ' Future
            ElseIf rs!F_Lev = Cur_Flevel Then
                If LastFinalCompleted = True Then
                    rs!Status = 2 ' Completed
                Else
                    rs!Status = 1  ' Current
                End If
            Else
                If rs!Status <> 3 Then ' Completed
                    rs!Status = 2
                End If
                     
            End If

            rs.Update        ' Save changes.
            rs.FindNext Criteria ' Locate next record.
        Loop

    
    End If
    rs.Close
    

SetCurrentFinal_Exit:
    Exit Sub

SetCurrentFinal_Error:
    MsgBox ("An exception has occured in SetCurrentFinal: " & Error)
    GoTo SetCurrentFinal_Exit

End Sub

Function SetHeatFormat(Heat)

    If HeatFormat = "ABCD" Then
        SetHeatFormat = Chr(64 + Heat)
    Else
        SetHeatFormat = Heat
    End If
    
End Function

Function SetSexFormat(Sex)

    If SexFormat = "Boys/Girls" Then
        If Sex = "F" Then
            SetSexFormat = "Girls"
        Else
            SetSexFormat = "Boys"
        End If
    Else
        If Sex = "F" Then
            SetSexFormat = "Female"
        Else
            SetSexFormat = "Male"
        End If
    End If
        
End Function


Sub Update_Lane_Assignments(E_Code, F_Lev, Heat)

    Dim Criteria As String, db As Database, rs As Recordset, LRS As Recordset
    Dim H_ID As Variant

    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset("CompEvents", DB_OPEN_DYNASET)   ' Create Recordset.
    
    Criteria = "[E_Code]=" & E_Code & " AND [F_Lev]=" & F_Lev & " AND [Heat]=" & Heat
    rs.FindFirst Criteria
    While Not rs.NoMatch

        If rs!Lane = 0 Then
            H_ID = DLookup("[H_ID]", "Competitors", "[PIN]=" & rs!PIN)
            
            rs.Edit
            rs!Lane = Calculate_Competitor_Lane(E_Code, F_Lev, H_ID, Heat)
            rs.Update
            
        End If
        
        rs.FindNext Criteria
        
    Wend

    rs.Close

End Sub

'*****************************************************************************************************************************
'Purpose:       -
'Parameters:    None
'Returns:       None
'Created By:    Andrew Rogers
'Created On:    Sun 16/Feb/2003
'Comments:      None
'*****************************************************************************************************************************
Public Sub UpdateEventCompetitorAge()
On Error GoTo UpdateEventCompetitorAge_Err

  Dim CArs As Recordset       ' Competitor Age
  Dim EArs As Recordset       ' Event Age
  Dim CEArs As Recordset      ' CompetiotrEventAge
  
  If SportsViewModule Then Exit Sub
  
  PleaseWaitMsg = "Updating competitor age information..."
  DoCmd.RunMacro "ShowPleaseWait"
  
  'Add to the CompetitorEventAge table all EventAge and Competitor age pairs
  
  Q = "UPDATE [CompetitorEventAge] SET [Tag]=TRUE"
  CurrentDb.Execute Q
  
  Set CArs = CurrentDb.OpenRecordset("SELECT Competitors.Age FROM Competitors GROUP BY Competitors.Age")
  Set EArs = CurrentDb.OpenRecordset("SELECT Events.Age FROM Events GROUP BY Events.Age")
  Set CEArs = CurrentDb.OpenRecordset("SELECT * FROM CompetitorEventAge", dbOpenDynaset)
  
  If EArs.BOF Then Exit Sub
  
  Do Until CArs.BOF Or CArs.EOF
    EArs.MoveFirst
    Do Until EArs.EOF
      If CompAgeSatisfiesEventAge(CArs!Age, EArs!Age) Then
        CEArs.FindFirst "[Cage]=" & CArs!Age & " AND [Eage]=""" & EArs!Age & """"
        If CEArs.NoMatch Then
          CEArs.AddNew
          CEArs!Cage = CArs!Age
          CEArs!Eage = EArs!Age
        Else
          CEArs.Edit
        End If
        CEArs!Tag = False
        CEArs.Update
        Debug.Print CArs!Age, EArs!Age
      End If
      EArs.MoveNext
    Loop
    CArs.MoveNext
  Loop
  
  Q = "DELETE * FROM CompetitorEventAge WHERE [Tag]=TRUE"
  CurrentDb.Execute Q
  

  
UpdateEventCompetitorAge_Exit:
  On Error Resume Next
  DoCmd.RunMacro "ClosePleaseWait"
  Exit Sub

UpdateEventCompetitorAge_Err:
  Call DisplayErrMsg("UpdateEventCompetitorAge")
  Resume UpdateEventCompetitorAge_Exit

End Sub

Private Function CompAgeSatisfiesEventAge(Cage As Byte, Eage As String) As Boolean
On Error GoTo CompAgeSatisfiesEventAge_Err
  
  CompAgeSatisfiesEventAge = False
  
  If Right(Eage, 2) = "_O" Then
    If Cage >= Val(Eage) Then CompAgeSatisfiesEventAge = True
    
  ElseIf Right(Eage, 2) = "_U" Then
    If Cage <= Val(Eage) Then CompAgeSatisfiesEventAge = True
  
  ElseIf Eage = "OPEN" Then
    CompAgeSatisfiesEventAge = True
  Else
    If Cage = Val(Eage) Then CompAgeSatisfiesEventAge = True
  End If
  
CompAgeSatisfiesEventAge_Exit:
  On Error Resume Next
  Exit Function

CompAgeSatisfiesEventAge_Err:
  Call DisplayErrMsg("CompAgeSatisfiesEventAge")
  Resume CompAgeSatisfiesEventAge_Exit

End Function

Sub UpdateEventCompetitorAgeOLD()
On Error GoTo UpdateEventCompetitorAge_Err
    
  If SportsViewModule Then Exit Sub
  
  PleaseWaitMsg = "Updating competitor age information..."
  DoCmd.RunMacro "ShowPleaseWait"
  
  'Stop
  
  Dim Criteria As String, db As Database, rs As Recordset, Q As Variant, i As Variant
  Dim Cage As Variant, Eage As Variant, CEArs As Recordset
  
  Q = "SELECT DISTINCT Competitors.Age FROM Competitors"
  
  Set rs = CurrentDb.OpenRecordset(Q, DB_OPEN_DYNASET)   ' Create dynaset.
  Set CEArs = CurrentDb.OpenRecordset("CompetitorEventAge", dbOpenDynaset)
  
  Do Until CEArs.BOF Or CEArs.EOF
    CEArs.Edit
    CEArs!Flag = True
    CEArs.Update
    CEArs.MoveNext
  Loop
  
  i = 0
  DoCmd.SetWarnings False
  DoCmd.RunSQL "delete * from CompetitorEventAge"
  DoCmd.SetWarnings True
  
  Do Until rs.EOF  ' Loop until no matching records.
    'Stop
    Cage = rs![Age]
    If Not IsNull(Cage) Then
      Eage = DetermineEventAge(Cage)
             
      CEArs.FindFirst "[Cage]=" & rs!Age & " AND [Eage]=""" & Eage & """"
      If CEArs.NoMatch Then
        With CEArs
          .AddNew
          !Cage = rs!Age
          !Eage = Eage
          !Flag = False
          .Update
        End With
      Else
        CEArs.Edit
        CEArs!Flag = False
        CEArs.Update
      End If
    End If
    rs.MoveNext
  Loop
  
  Q = "DELETE CompetitorEventAge.*, CompetitorEventAge.Flag "
  Q = Q & "FROM CompetitorEventAge "
  Q = Q & "WHERE Flag=True"

  CurrentDb.Execute Q
  
  rs.Close
  CEArs.Close
  
  DoCmd.RunMacro "ClosePleaseWait"

UpdateEventCompetitorAge_Exit:
    Exit Sub

UpdateEventCompetitorAge_Err:
    MsgBox "UpdateEventCompetitorAge:" & Error$, vbCritical
    GoTo UpdateEventCompetitorAge_Exit


End Sub

Function Work_AutoEventNumber()

    Dim Criteria As String, db As Database, rs As Recordset, x As Variant
    
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset("Work-Heats in Some Order", DB_OPEN_DYNASET)   ' Create Recordset.
    
    x = 1
    rs.MoveFirst
    While Not rs.EOF
        rs.Edit
        rs!E_Number = x
        rs.Update
        x = x + 1
        rs.MoveNext
    Wend

    rs.Close


End Function
Function AddAppProperty(strName As String, varType As Variant, varValue As Variant) As Integer
  Dim dbs As Database, prp As Property
  Const conPropNotFoundError = 3270

  Set dbs = CurrentDb
  On Error GoTo AddProp_Err
  dbs.Properties(strName) = varValue

AddAppProperty = True

AddProp_Bye:
  Exit Function

AddProp_Err:
  If Err = conPropNotFoundError Then
    Set prp = dbs.CreateProperty(strName, varType, varValue)
    dbs.Properties.Append prp
    Resume
  Else
    AddAppProperty = False
    Resume AddProp_Bye
  End If
End Function

Public Function ConvertNullToZero(V As Variant)

  If IsNull(V) Then
    ConvertNullToZero = 0
  Else
    ConvertNullToZero = V
  End If
  
End Function

Public Sub QuitSportsAdministrator(F As Form)

  DoCmd.Close acForm, F.Name
  Application.Quit
  
End Sub

Public Function NoFormRecords(rstForm As Recordset) As Boolean

  If rstForm.RecordCount = 0 Then
    NoFormRecords = True
  Else
    NoFormRecords = False
  End If

End Function

' *******************************************************
' *** Check if variable is Empty
' *******************************************************
Public Function VarEmpty(V As Variant) As Boolean
  
'  Stop
  
  If IsNull(V) Then
    VarEmpty = True
  ElseIf IsNumeric(V) Then
    If V = 0 Then VarEmpty = True
  ElseIf Trim(V) = "" Then
    VarEmpty = True
  Else
    VarEmpty = False
  End If
  
End Function

Private Function PopUpFormsVisible(Visibility As Boolean)
On Error GoTo PopUpFormsVisible_Err

  Dim F As Form
  For Each F In Forms
    If F.PopUp Then
      If Visibility = False Then 'Hide All Popup forms
        If F.Visible Then
          F.Visible = False
          F.Tag = "Hidden By PopUpFormsVisible"
        End If
        
      Else ' SHow all popup forms
        If F.Tag = "Hidden By PopUpFormsVisible" Then
          F.Visible = True
          F.Tag = ""
        End If
      End If
    End If
  Next

  DoEvents

PopUpFormsVisible_Exit:
  Exit Function
  
PopUpFormsVisible_Err:
  MsgBox "An error has occurred in [PopUpFormsVisible]: " & Err.Description, vbCritical
  Resume PopUpFormsVisible_Exit
  
End Function

Public Function DisplayPrintDialog()
On Error GoTo DisplayPrintDialog_Err

  Dim ObType As Variant, ClosedOpenReportsForm As Boolean
  
  Call PopUpFormsVisible(False)
  
  ObType = Application.CurrentObjectType
  
  DoCmd.SelectObject Application.CurrentObjectType, Application.CurrentObjectName
  DoCmd.RunCommand acCmdPrint
  
  'If ObType = acForm Then
  '
  '  DoCmd.PrintOut
  'ElseIf ObType = acReport Then
  '  DoCmd.OpenReport Application.CurrentObjectName, acViewNormal
  'End If
  
  Call PopUpFormsVisible(True)

DisplayPrintDialog_Exit:
  Exit Function
  
DisplayPrintDialog_Err:
  If Err.Number <> 2212 Then ' Print cancelled
    MsgBox "An error has occurred in [DisplayPrintDialog]: " & Err.Number & " - " & Err.Description, vbCritical
  End If
  On Error Resume Next
  Call PopUpFormsVisible(True)
  
  GoTo DisplayPrintDialog_Exit
  
End Function


Public Function SportsViewModule() As Boolean
On Error GoTo SportsViewModule_Err

  If Right(CurrentDb.Name, 14) = "sportsview.mdb" Or Right(CurrentDb.Name, 14) = "sportsview.mde" _
  Or Right(CurrentDb.Name, 16) = "sportsview.accdb" Then
    SportsViewModule = True
  Else
    SportsViewModule = False
  End If
  
SportsViewModule_Exit:
  Exit Function
  
SportsViewModule_Err:
  MsgBox Err.Description
  
End Function

Public Function LFCR()
    LFCR = Chr(13) & Chr(10)
End Function