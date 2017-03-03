Option Compare Database

Public Function AddTrustedLocation(trustPath As String, Optional trustName As String = "") As Integer
On Error GoTo err_proc
'WARNING:  THIS CODE MODIFIES THE REGISTRY
'sets registry key for 'trusted location - This code was obtained from http://allenbrowne.com/Access2007.html
' This is requried as TransferDatabase will give many wanring prompts if data file is not in Trusted Location

  Dim intLocns As Integer
  Dim i As Integer
  Dim intNotUsed As Integer
  Dim strLnKey As String
  Dim reg As Object
  Dim strPath As String
  Dim strTitle As String
  AddTrustedLocation = -1
  
  Set reg = CreateObject("Wscript.Shell")
  
  If (DirExists(trustPath)) Then
    strPath = trustPath
  Else
      AddTrustedLocation = -1
      Exit Function
  End If
  
  If (trustName = "") Then
    strTitle = Application.CurrentProject.Name
  Else
    strTitle = trustName
  End If
  
  'Specify the registry trusted locations path for the version of Access used
  strLnKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Format(Application.Version, "##,##0.0") & _
             "\Access\Security\Trusted Locations\Location"
On Error GoTo err_proc0
  'find top of range of trusted locations references in registry
  For i = 999 To 0 Step -1
      reg.RegRead strLnKey & i & "\Path"
      GoTo chckRegPths        'Reg.RegRead successful, location exists > check for path in all locations 0 - i.
checknext:
  Next
  MsgBox "Unexpected Error - No Registry Locations found", vbExclamation
  GoTo exit_proc
  
  
chckRegPths:
'Check if Currentdb path already a trusted location
'reg.RegRead fails before intlocns = i then the registry location is unused and
'will be used for new trusted location if path not already in registy
On Error GoTo err_proc1:
  For intLocns = 1 To i
      reg.RegRead strLnKey & intLocns & "\Path"
      'If Path already in registry -> exit
      If InStr(1, reg.RegRead(strLnKey & intLocns & "\Path"), strPath) = 1 Then GoTo exit_proc
      
NextLocn:
  Next
  
  If intLocns = 999 Then
      MsgBox "Location count exceeded - unable to write trusted location to registry", vbInformation, strTitle
      GoTo exit_proc
  End If
  'if no unused location found then set new location for path
  If intNotUsed = 0 Then intNotUsed = i + 1
  
'Write Trusted Location regstry key to unused location in registry
On Error GoTo err_proc:
  strLnKey = strLnKey & intNotUsed & "\"
  reg.RegWrite strLnKey & "AllowSubfolders", 1, "REG_DWORD"
  reg.RegWrite strLnKey & "Date", Now(), "REG_SZ"
  reg.RegWrite strLnKey & "Description", strTitle, "REG_SZ"
  reg.RegWrite strLnKey & "Path", strPath & "\", "REG_SZ"
  
  AddTrustedLocation = intNotUsed
  
exit_proc:
  Set reg = Nothing
  Exit Function
  
err_proc0:
  Resume checknext
  
err_proc1:
  If intNotUsed = 0 Then intNotUsed = intLocns
  Resume NextLocn
err_proc:
  MsgBox Err.Description, , strTitle
  Resume exit_proc
  
End Function


Public Sub DelTrustedLocation(locationNum As Integer)
    ' Sub to remove an added trusted location
    ' Uses the number for location<n>
    
    'On Error Resume Next
    
    Dim objShell As Object
    Dim strKey As String
    
    If (locationNum < 0) Then
        Exit Sub
    End If
    
    Set objShell = CreateObject("Wscript.Shell")
    
    strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Format(Application.Version, "##,##0.0") & _
             "\Access\Security\Trusted Locations\Location" & locationNum & "\"
             
    objShell.RegDelete strKey


End Sub