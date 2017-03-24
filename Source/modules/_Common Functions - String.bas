Option Compare Database
Option Explicit

' ########################################################
' ### STRING FUNCTIONS
' ########################################################


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

' Extract the n-th item from a string list seperated by delimiter
Public Function StringParse(s As String, ItemNum As Byte, Optional delimiter As String = "|") As String
On Error GoTo StringParse_Err
  
  'Retrieves the String Item at Item Number
  
  Dim L As Integer, CurrentItem As Long, c As String, PS As String, Complete As Boolean, i As Long
  
  s = s & delimiter
  
  L = Len(s)
  
  CurrentItem = 0
  
  If IsMissing(delimiter) Then delimiter = "|"
  PS = ""
  Complete = False
  i = 0
  Do Until i > L Or Complete
    i = i + 1
    c = Mid(s, i, 1)
    If c = delimiter Then
      CurrentItem = CurrentItem + 1
      If CurrentItem = ItemNum Then
        Complete = True
      Else
        PS = ""
      End If
      
    Else
      PS = PS & c
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