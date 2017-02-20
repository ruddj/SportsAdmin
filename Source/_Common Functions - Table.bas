Option Compare Database
Option Explicit

' *******************************************************
' *** OPENS A LINKED TABLE FOR USE WITH THE SEEK COMMAND
' *******************************************************

'This code was originally written by Michel Walsh. It is not to be altered or distributed,
'except as part of an application. You are free to use it in any application, provided the copyright notice is left unchanged.
'Code Courtesy of Michel Walsh

Public Function OpenForSeek(TableName As String, Optional Quiet, Optional IsQuery) As Recordset
On Error GoTo OpenForSeek_Err

  If IsMissing(Quiet) Then Quiet = False
' Assume MS-ACCESS table
  If CurrentDb().TableDefs(TableName).Connect = "" Then
    If Not Quiet Then MsgBox "Opening a LOCAL table for seek."
    Set OpenForSeek = CurrentDb.OpenRecordset(TableName, dbOpenTable)
  Else
    Set OpenForSeek = DBEngine.Workspaces(0).OpenDatabase _
                    (Mid(CurrentDb().TableDefs(TableName).Connect, _
                    11), False, False, "").OpenRecordset(TableName, _
                    dbOpenTable)
  End If
  
OpenForSeek_Exit:
  Exit Function
  
OpenForSeek_Err:
  Stop
  MsgBox "An error has occurred in [OpenForSeek]: " & Err.Description, vbCritical
  
End Function