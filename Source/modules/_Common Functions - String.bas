Option Compare Database
Option Explicit

' ########################################################
' ### STRING FUNCTIONS
' ########################################################

'Public Function NotEmptyString(v As String)
'Public Sub AddToFilter(ByRef Filter As String, Criteria As String, Operator As String)
'Public Function ExtractDirectory(F)

' *******************************************************
' *** DETERMINES IF STRING IS EMPTY (EITHER NULL OR "")
' *******************************************************
Public Function NotEmptyString(V As String)

  If IsNull(V) Then
    NotEmptyString = False
  ElseIf Trim(V) = "" Then
    NotEmptyString = False
  Else
    NotEmptyString = True
  End If
  
End Function

Public Sub AddToFilter(ByRef Filter As String, Criteria As String, Operator As String)

  If Filter <> "" Then
    Filter = Filter & " " & Operator & " " & Criteria
  Else
    Filter = Criteria
  End If

End Sub

Public Function ExtractFileName(s) As String
  ExtractFileName = ExtractFileNameFromPath(s)
End Function

Public Function ExtractFileNameFromPath(s) As String
  Dim L As Integer
  Dim i As Integer
  Dim fn As String
  
  fn = ""
  i = Len(s)
  
  Do Until (Mid(s, i, 1) = "\" Or i = 0)
    fn = Mid(s, i, 1) & fn
    i = i - 1
  Loop

  ExtractFileNameFromPath = fn
  
End Function


Public Function CRLF(Count As Integer) As String

  Dim s  As String, i As Integer
    
  s = ""
  For i = 1 To Count
    s = s & vbCr & vbLf
  Next
  
  CRLF = s
  
End Function
'*************************************************************************************
'  Check if an item is in a group
'*************************************************************************************
Public Function InGroup(Item, iGroup As String) As Boolean
On Error GoTo InGroup_Err

  'Format of iGroup is 'Item1,Item2,Item3 ...'
  
  Dim iL As Integer, iG As Integer
  Dim gLen As Integer, x As Integer, C As String, gItem As String
  Dim CreatingGroupItem As Boolean, Finish As Boolean
  
  iL = Len(Item)
  gLen = Len(iGroup)
  x = 1
  CreatingGroupItem = True
  InGroup = False
  Finish = False
  gItem = ""
  
  Do Until InGroup Or Finish
    
    If x > gLen Then
      Finish = True
      CreatingGroupItem = False
    End If
      
    If Not CreatingGroupItem Then
      If Trim(Item) = Trim(gItem) Then InGroup = True
      CreatingGroupItem = True
      gItem = ""
      x = x + 1
    Else
      C = Mid(iGroup, x, 1)
      If C = "," Then
        CreatingGroupItem = False
      Else
        CreatingGroupItem = True
        gItem = gItem & C
        x = x + 1
      End If
    End If
  Loop
  
InGroup_Exit:
  Exit Function
  
InGroup_Err:
  Call DisplayErrMsg("InGroup")
  
End Function

Public Function ConvertToTitleCase(s As String, Optional IsName) As String
On Error GoTo ConvertToTitleCase_Err

  Dim i As Integer, PrevChar As String, Uppercase As Boolean, NewS As String, C As String
  Dim NewWordLetters As Integer
  NewS = ""
  PrevChar = ""
  Uppercase = True
  NewWordLetters = 0
  If IsMissing(IsName) Then IsName = False
  
  For i = 1 To Len(s)
    NewWordLetters = NewWordLetters + 1
    C = Mid(s, i, 1)
    If Uppercase Then
      NewS = NewS & UCase(C)
      Uppercase = False
    Else
      NewS = NewS & LCase(C)
    End If
    If Not ((Asc(C) >= 65 And Asc(C) <= 90) Or (Asc(C) >= 97 And Asc(C) <= 122)) Then
      Uppercase = True
      NewWordLetters = 0
    End If
    
    If IsName And NewWordLetters = 2 Then
      If Mid(s, i - 1, 2) = "mc" Then Uppercase = True
    End If
    
  Next i
    
  ConvertToTitleCase = NewS

ConvertToTitleCase_Exit:
  Exit Function
  
ConvertToTitleCase_Err:
  MsgBox "An error has occurred in [ConvertToTitleCase]: " & Err.Description, vbCritical
  
End Function

Public Function ConvertToUpperLowerCase(s As String) As String
On Error GoTo ConvertToUpperLowerCase_Err

  Dim i As Integer, PrevChar As String, Uppercase As Boolean, NewS As String, C As String
  Dim FirstName As Boolean
  
  NewS = ""
  Uppercase = True
  FirstName = True
  
  For i = 1 To Len(s)
    C = Mid(s, i, 1)
    If Uppercase Then
      NewS = NewS & UCase(C)
    Else
      NewS = NewS & LCase(C)
    End If
    
    If FirstName Then
      Uppercase = True
    Else
      Uppercase = False
    End If
    
    If Not ((Asc(C) >= 65 And Asc(C) <= 90) Or (Asc(C) >= 97 And Asc(C) <= 122)) Then
      Uppercase = True
      FirstName = False
    End If
    
  Next i
    
  Debug.Print NewS
  
  ConvertToUpperLowerCase = NewS

ConvertToUpperLowerCase_Exit:
  Exit Function
  
ConvertToUpperLowerCase_Err:
  MsgBox "An error has occurred in [ConvertToUpperLowerCase]: " & Err.Description, vbCritical
  
End Function

Public Function strReplace(FullString As String, SearchString As String, ReplaceString As String) As String
On Error GoTo strReplace_Err

  Dim NewString As String, SearchStringLength As Integer, i As Integer
  
  NewString = ""
  
  SearchStringLength = Len(SearchString)
  For i = 1 To Len(FullString) '(Len(FullString) - SearchStringLength)
    If Mid(FullString, i, SearchStringLength) = SearchString Then
      NewString = NewString & ReplaceString
      i = i + SearchStringLength - 1
    Else
      NewString = NewString & Mid(FullString, i, 1)
    End If
  Next i
  
  strReplace = NewString
  
strReplace_Exit:
  Exit Function
  
strReplace_Err:
  MsgBox "An error has occurred in [strReplace]: " & Err.Description, vbCritical

End Function

Private Sub test()
  Dim Q As String
  
  Q = "Andrew" & vbCr & vbLf & "Rogers"
  Debug.Print "|" & Q & "|"
  
  Debug.Print "*" & strReplace(Q, vbCr & vbLf, ", ") & "*"
  
End Sub


'*****************************************************************************************************************************
'Purpose:       Checks for one string in another
'Parameters:    None
'Returns:       None
'Created By:    Andrew Rogers
'Created On:    Thu 11/Jul/2002
'Comments:      None
'*****************************************************************************************************************************
Public Function strIn(FullString As String, SearchString As String, Optional ByRef Position) As Boolean
On Error GoTo strIn_Err

  Dim SearchStringLength As Integer, i As Integer
  
  strIn = False
  
  SearchStringLength = Len(SearchString)
  For i = 1 To Len(FullString) '(Len(FullString) - SearchStringLength)
    If Mid(FullString, i, SearchStringLength) = SearchString Then
      strIn = True
      Exit For
    End If
  Next i
  
  If Not IsMissing(Position) Then Position = i
  
strIn_Exit:
  On Error Resume Next
  Exit Function

strIn_Err:
  Call DisplayErrMsg("strIn")
  Resume strIn_Exit

End Function


Public Sub DateQuickEnter(D As TextBox)
On Error GoTo DateQuickEnter_Err

  If IsNumeric(D.Text) Then
    If Len(D.Text) = 6 Then
      Dim NewText As String, i As Integer
      NewText = ""
      
      For i = 1 To Len(D.Text)
        
        NewText = NewText & Mid(D.Text, i, 1)
        If i Mod 2 = 0 And i <> 6 Then NewText = NewText & "/"
      Next i
    
      If IsDate(NewText) Then D.Text = NewText
    End If
  End If
  
DateQuickEnter_Exit:
  Exit Sub
  
DateQuickEnter_Err:
  Call DisplayErrMsg("DateQuickEnter")
  Resume DateQuickEnter_Exit
  
End Sub


Public Function RemoveDoubleSpaces(ByRef sn As String)
  
  Call RemoveDoubleCharacters(sn, " ")

End Function

Public Function RemoveDoubleCharacters(ByRef sn As String, Char As String)
  
  Dim sCount As Integer, NewSN As String, i As Integer
  
  sCount = 0
  For i = 1 To Len(sn)
    If Mid(sn, i, 1) <> Char Then
      sCount = 0
    Else
      sCount = sCount + 1
    End If
    If sCount <= 1 Then
      NewSN = NewSN & Mid(sn, i, 1)
    End If
  Next i
  
  sn = NewSN
  
End Function

Public Function ReplaceCharacter(sn As String, Char As String, NewChar As String) As String
  
  Dim sCount As Integer, NewSN As String, i As Integer
  
  sCount = 0
  For i = 1 To Len(sn)
    If Mid(sn, i, 1) = Char Then
      NewSN = NewSN & NewChar
    Else
      NewSN = NewSN & Mid(sn, i, 1)
    End If
  Next i
  
  ReplaceCharacter = NewSN

End Function


Public Function StringParse(s As String, ItemNum As Byte, Optional Delimiter As String = "|") As String
On Error GoTo StringParse_Err
  
  'Retrieves the String Item at Item Number
  
  Dim L As Integer, CurrentItem As Long, C As String, PS As String, Complete As Boolean, i As Long
  
  s = s & Delimiter
  
  L = Len(s)
  
  CurrentItem = 0
  
  If IsMissing(Delimiter) Then Delimiter = "|"
  PS = ""
  Complete = False
  i = 0
  Do Until i > L Or Complete
    i = i + 1
    C = Mid(s, i, 1)
    If C = Delimiter Then
      CurrentItem = CurrentItem + 1
      If CurrentItem = ItemNum Then
        Complete = True
      Else
        PS = ""
      End If
      
    Else
      PS = PS & C
    End If
  Loop
  
  If Complete Then
    StringParse = Trim(PS)
  Else
    StringParse = ""
  End If
  
StringParse_Exit:
  Exit Function
  
StringParse_Err:
  Call DisplayErrMsg("StringParse")
  Resume StringParse_Exit
  
End Function

Public Function ExtractValue(VarName As String, StringToParse As Variant, Optional Delimiter As String = ";") As Variant
On Error GoTo ExtractValue_Err

  'Search for variable VarName.  If not found return null
  ' Format: VarName:=VALUE; VarName2:=VALUE2
  
  Dim s As String
  Dim L As Integer
  Dim V As String
  Dim N As String
  Dim i As Long
  Dim Found As Boolean
  Dim GettingValue As Boolean
  Dim Char As String
  
  If IsNull(StringToParse) Then
    ExtractValue = Null
    Exit Function
  End If
  
  s = Delimiter & StringToParse & Delimiter
    
  i = 1
  L = Len(s)
  Found = False
  GettingValue = False
  Do Until i > L Or Found
    Char = Mid(s, i, 1)
    
    If Char = Delimiter Then
      If Trim(N) = Trim(VarName) Then
        Found = True
      Else
        N = ""
        V = ""
        GettingValue = False
      End If
    Else
      If Char = ":" Then
        If Mid(s, i + 1, 1) = "=" Then
          GettingValue = True
          i = i + 1
        End If
      Else
        If GettingValue Then
          V = V & Char
        Else
          N = N & Char
        End If
      End If
    End If
    i = i + 1
  Loop
  
  If Found Then
    V = Trim(V)
    If IsNumeric(V) Then
      ExtractValue = Val(V)
      
    ElseIf IsDate(V) Then
      ExtractValue = CDate(V)
      
    Else
      ExtractValue = V
    End If
  Else
    ExtractValue = Null
  End If
  
ExtractValue_Exit:
  Exit Function
  
ExtractValue_Err:
  Call DisplayErrMsg("ExtractValue")
  Resume ExtractValue_Exit
  
End Function


'*****************************************************************************************************************************
'Purpose:       Converts UK dates to US format by using Text for month
'Parameters:    None
'Returns:       None
'Created By:    Andrew Rogers
'Created On:    Tue 17/Sep/2002
'Comments:      None
'*****************************************************************************************************************************
Public Function DateUKtoUS(D As String) As String
On Error GoTo DateUKtoUS_Err

  Dim vD As Variant, vM As Variant, vY As Variant
  Dim iPart As Byte, C As String
  Dim i As Integer
  
  vD = Null
  vM = Null
  vY = Null
  iPart = 0
  
  For i = 1 To Len(D)
    C = Mid(D, i, 1)
    If C = "\" Or C = "/" Or C = "-" Then
      iPart = iPart + 1
    Else
      Select Case iPart
      Case 0
        vD = vD & C
      Case 1
        vM = vM & C
      Case 2
        vY = vY & C
        
      End Select
    End If
  Next i
  
  If IsNumeric(vM) Then
    Select Case vM
    Case 1
      vM = "Jan"
    Case 2
      vM = "Feb"
    Case 3
      vM = "Mar"
    Case 4
      vM = "Apr"
    Case 5
      vM = "May"
    Case 6
      vM = "Jun"
    Case 7
      vM = "Jul"
    Case 8
      vM = "Aug"
    Case 9
      vM = "Sep"
    Case 10
      vM = "Oct"
    Case 11
      vM = "Nov"
    Case 12
      vM = "Dec"
    End Select
  End If
  
  DateUKtoUS = vD & "/" & Nz(vM, Month(Now)) & "/" & Nz(vY, Year(Now))
  
DateUKtoUS_Exit:
  On Error Resume Next
  Exit Function

DateUKtoUS_Err:
  Call DisplayErrMsg("DateUKtoUS")
  Resume DateUKtoUS_Exit

End Function