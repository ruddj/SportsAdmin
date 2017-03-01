Option Compare Database
Option Explicit


Public Sub BackupCurrentCarnival(SameAsCarnival As Boolean, BackupPath As Variant)
On Error GoTo BackupCurrentCarnival_Err

  Dim Db As Database, T As TableDef, OrigFile As Variant, NewFile As Variant, Q As String
  
  Call CloseAlwaysOpenRS
  
  Set Db = CurrentDb
  Set T = Db.TableDefs("Competitors")
  OrigFile = T.connect
  OrigFile = Right(OrigFile, Len(OrigFile) - 10)
  
  If SameAsCarnival Then ' Backup to the same folder as the carnival
    If StrConv(Right(OrigFile, 6), vbLowerCase) = ".accdb" Then
      NewFile = Left(OrigFile, Len(OrigFile) - 6) & "_backup.accdb"
    Else
      NewFile = Left(OrigFile, Len(OrigFile) - 4) & "_backup.mdb"
    End If
  ElseIf IsNull(BackupPath) Then
    MsgBox ("You must specify a folder to backup the carnival into.  This can be done in the Utilities form.")
    Exit Sub
  Else
    If Right(BackupPath, 1) <> "\" Then BackupPath = BackupPath & "\"
    NewFile = BackupPath & GetCarnivalFile(OrigFile)
  End If
  If FileExists(NewFile) Then Kill NewFile
    
  DBEngine.CompactDatabase OrigFile, NewFile
  Response = MsgBox("The Carnival has been backed up to " & NewFile, vbInformation)
  
BackupCurrentCarnival_Exit:
  Call OpenAlwaysOpenRS
  Exit Sub
  
BackupCurrentCarnival_Err:
  
  If Err.Number = 3356 Then 'Carnival file open by someone else
    Q = "Another user appears to have the Sports Administrator open and is working on this carnival file.  "
    Q = Q & "Backups cannot be performed when there is more than one user working on the carnival."
    MsgBox Q, vbExclamation
  Else
    MsgBox "An error has occured in [BackupCurrentCarnival]: " & Err.Description, vbCritical
  End If
  GoTo BackupCurrentCarnival_Exit

End Sub