Option Compare Database
Option Explicit

'************** Code Start **************
'This code was originally written by Terry Kreft.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code courtesy of
'Terry Kreft

Private Type BROWSEINFO
  hOwner As LongPtr
  pidlRoot As LongPtr
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As LongPtr
  lParam As LongPtr
  iImage As Long
End Type

Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias _
            "SHGetPathFromIDListA" (ByVal pidl As LongPtr, _
            ByVal pszPath As String) As Boolean
            
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias _
            "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
            As LongPtr
            
Private Const BIF_RETURNONLYFSDIRS = &H1

Public Function BrowseFolder(szDialogTitle As String) As String
  Dim X As Boolean, bi As BROWSEINFO, dwIList As LongPtr
  Dim szPath As String, wPos As Integer
  
    With bi
        .hOwner = hWndAccessApp
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If
End Function

'*********** Code End *****************