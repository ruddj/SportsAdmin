Option Compare Database
Option Explicit

Private Sub displayerrs()

  Dim lngCode As Long
  Dim strAccessErr As String
  Const conAppObjectError = "Application-defined or object-defined error"

  
  DoCmd.Hourglass True

  
  For lngCode = 3000 To 3200
    On Error Resume Next
    ' Raise each error.
    strAccessErr = AccessError(lngCode)
    ' Skip error numbers without associated strings.
    If strAccessErr <> "" Then

    ' Skip codes that generate application or object-defined errors.
      If strAccessErr <> conAppObjectError Then
        Debug.Print "Code: " & lngCode & "  Err: " & strAccessErr
      End If
    End If
  Next lngCode
  
  DoCmd.Hourglass False
  
End Sub

Public Sub DisplayErrMsg(CurrentProcedure, Optional Var1, Optional Var2)

'  Stop
  
  If IsMissing(Var1) Then Var1 = ""
  If IsMissing(Var2) Then Var2 = ""
  
  Dim DialogImage As Variant, Title As String
  
  DialogImage = vbCritical
  
  Select Case Err.Number
      
    Case 9    '########## Subscript out of range - Usually occurs on TreeView menus when developing
      If CurrentUser = "Owner" Then
        Q = ""
      Else
        Q = "Subscript out of range"
      End If
    
    Case 2046  '########## Tried to delete a non-existant record
      Q = ""
    
    Case 2212  '########## An error occured trying to print a report
      Q = "An error occured trying to print the report."
      DialogImage = vbInformation
    
    Case 2501  '########## A report was opened with no data and the open event was then cancelled so dont show error msg
      Q = ""
  
    Case 2603 '########## No permissions to open form or report
      Q = "You do not have permission to view this form or report.  This error should not normally occur in SchoolPRO v2.0 and later."
      DialogImage = vbExclamation
       
    Case 3107, 3108, 3109, 3111 '########## No permissions to modify table
      Q = "You do not have sufficient permissions to modify records in this table."
      DialogImage = vbInformation
       
    Case 3078 '########### ???
      Select Case CurrentProcedure
        Case "OpenAllLinkedDatabases"
          Q = "The table " & Var1 & " is not present in " & Var2 & ".  Database functionality will remain the same however a decrease in performance may be experienced."
          GoSub ShowErrMsg
          Resume Next
          
        Case Else
          GoSub UseStandardErrorMsg
      End Select
    
    Case 3112 '########## Insuffucent permissions to view table
      Select Case CurrentProcedure
                  
        Case "OpenAllLinkedDatabases", "CheckIfLoginAllowed"
          Q = "You do not have the necessary permissions to view the table '" & Var1 & "'. " & CRLF(2)
          Q = Q & "All users should have read permissions for this table.  Please talk to the SchoolPRO administrator "
          Q = Q & "to resolve the issue."
          Title = "INSUFFICIENT PERMISSIONS"
          DialogImage = vbExclamation
          
        Case Else
          Q = "You do not have permission to view the table '" & Var1 & "'.  " & CRLF(2)
          Q = Q & "You should not use this SchoolPRO module until this issue is resolved by the SchoolPRO Administrator."
          Title = "INSUFFICIENT PERMISSIONS"
          DialogImage = vbExclamation
      End Select
      
    Case 64519 '########## Tried to do something with a deleted record
      Q = "The record has already been deleted."
      DialogImage = vbInformation
      
    Case Else '########## Error that is not specifically catered for
      GoSub UseStandardErrorMsg
      
  End Select
    
  GoSub ShowErrMsg
  
DisplayErrMsg_Exit:
  Exit Sub
  
DisplayErrMsg_err:
  MsgBox "An error has occurred in [DisplayErrMsg]: " & Err.Description
  Resume DisplayErrMsg_Exit
  
'-------------------
UseStandardErrorMsg:
  Q = "An error has occurred in [" & CurrentProcedure & "]:" & CRLF(2)
  Q = Q & "Error: " & Err.Description & CRLF(2)
  Q = Q & "Error#: " & Err.Number
  Return
  
ShowErrMsg:
  If Q <> "" Then MsgBox Q, DialogImage, Title
   Return
End Sub
  