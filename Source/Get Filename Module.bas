
'-------------------------------------------------------------------------------
' MODULE: Get Filename Module
' Provides the following functions. See functions for parameter definitions
'
'   GetFileName     : Retrieve a filename
'   SelectDir       : Retrieve a directory
'   FileExists      : Check existence of a file
'   MatchFiles      : Retrives a list of files that match given driteria
'



    Option Compare Database                                 ' Use database order for string comparisons
    Option Explicit
    Global Const NoFileSelection = "No File Selected"       ' This value is returned when no file is selected
    Type FileName_Info
        hwndOwner As Integer
        szFilter As String * 255
        szCustomFilter As String * 255
        nFilterIndex As Long
        szFile As String * 255
        szFileTitle As String * 255
        szInitialDir As String * 255
        szTitle As String * 255
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        szDefExt As String * 255
    End Type
    'Declare Function Get_FileName Lib "MSAU200.DLL" Alias "#1" (gfni As FileName_Info, ByVal fOpen As Integer) As Long
    'Declare Function Get_Directory Lib "SWU2016.DLL" Alias "#2" (ByVal Hwnd As Integer, ByVal sTitle As String, ByVal sDir As String) As Integer

Function ConvertNull(ByVal Value As Variant, ByVal subs As Variant) As Variant
'------------------------------------------------------------------------------
' Checks value
' If NULL , returns subs
' else returns value
    
    ConvertNull = IIf(IsNull(Value), subs, Value)
End Function

Function FileExists(ByVal FileName As String) As Variant
'---------------------------------------------------------------
' Returns variant indicating files existence
' -1 (True) : The file exists
' 0 (False) : The file does not exist
' 71        : The disk drive is not ready (floppy drive)
' 68        : The device is unavailable (an unconnected drive)

    On Error GoTo AssignErrorCode
    If (InStr(FileName, "*") = 0) And (InStr(FileName, "?") = 0) Then
        FileExists = (Dir(FileName) <> "")
    Else
        GoTo AssignErrorCode
    End If
    Exit Function
AssignErrorCode:
    FileExists = False  ' was = Err
    Resume Next
End Function

Function xGetFileName(ByVal Title As String, ByVal FilterOptions As String, ByVal FilterIndex As Integer, ByVal FilterDefault As String) As String
'-------------------------------------------------------------------------------
'  Return the file name chosen by user in OpenFile dialog box.
'  (This function works in conjunction with GetMDBName2 and StringFromSz to
'  display a File-Open dialog that prompts user for location of external reference files.
'  It uses code found in WZLIB.MDA.)
' Input parameters:
'   Title           : title for the window
'   FilterOptions   : String in the form {Description | Filter | } |
'   FilterIndex     : points to default option eg 1
'   FilterDefault   : Default filter extension eg "*"

    On Error GoTo Err_GetFileName
    Const OFN_SHAREAWARE = &H4000
    Const OFN_PATHMUSTEXIST = &H800
    Const OFN_HIDEREADONLY = &H4
    Dim OFN As FileName_Info
    OFN.hwndOwner = 0                                           ' Fill ofn structure which is passed to wlib_GetFileName
    OFN.szFilter = FilterOptions                                ' OFN.szFilter = "Databases (*.mdb)|*.mdb|All(*.*)|*.*||"
    OFN.nFilterIndex = FilterIndex
    OFN.szTitle = Title
    OFN.Flags = OFN_SHAREAWARE Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY
    OFN.szDefExt = FilterDefault
    'If (GetFileName2(OFN, True) = False) Then                   ' Call wlib_GetFileName function and interpret results.
    '    GetFileName = StringFromSz(OFN.szFile)
    'Else
    '    GetFileName = NoFileSelection
    'End If
Exit_GetFileName:
    Exit Function
Err_GetFileName:
    MsgBox Error$
    Resume Exit_GetFileName
End Function

Function xGetFileName2Old(gfni As FileName_Info, ByVal fOpen As Integer) As Long
'-------------------------------------------------------------------------------
'  This function acts as a cover to MSAU_GetFileName in MSAU200.DLL.
'  wlib_GetFileName terminates all strings in gfni structure with nulls and
'  then calls DLL version of function.  Upon returning from MSAU200.DLL, null
'  characters are removed from strings in gfni.
    
    Dim lRet As Long
    gfni.szFilter = RTrim$(gfni.szFilter) & Chr$(0)
    gfni.szCustomFilter = RTrim$(gfni.szCustomFilter) & Chr$(0)
    gfni.szFile = RTrim$(gfni.szFile) & Chr$(0)
    gfni.szFileTitle = RTrim$(gfni.szFileTitle) & Chr$(0)
    gfni.szInitialDir = RTrim$(gfni.szInitialDir) & Chr$(0)
    gfni.szTitle = RTrim$(gfni.szTitle) & Chr$(0)
    gfni.szDefExt = RTrim$(gfni.szDefExt) & Chr$(0)
    'lRet = Get_FileName(gfni, fOpen)
    gfni.szFilter = StringFromSz(gfni.szFilter)
    gfni.szCustomFilter = StringFromSz(gfni.szCustomFilter)
    gfni.szFile = StringFromSz(gfni.szFile)
    gfni.szFileTitle = StringFromSz(gfni.szFileTitle)
    gfni.szInitialDir = StringFromSz(gfni.szInitialDir)
    gfni.szTitle = StringFromSz(gfni.szTitle)
    gfni.szDefExt = StringFromSz(gfni.szDefExt)
'    GetFileName2 = lRet

End Function

Function MatchFiles(ByVal HdlWnd As Integer, Criteria() As String, FileArray() As String) As Variant
'--------------------------------------------------------------------------------
' Returns a sorted array of filenames that match the given parameter "Criteria". The user is
' prompted for a directory in which the comparison is made.
'
' Parameters
'       HdlWnd      : The window handle of the calling form. Identified by Me.Hwnd
'       Criteria    : Array of strings in the form of a filename. E.G. "*.MDB","*.MDA"
'       FileArray   : Returned array  or  unmodified as signalled by function result
'

    On Error GoTo Err_MatchFiles
    Dim FileMatch As String, i As Integer, j As Integer, k As Integer, L As Integer, m As Integer
    Dim FirstMatch As Variant, Result As String, ArrayMax  As Integer, Completed As Variant
    Dim temp As Variant
    Completed = False
    FirstMatch = True
    L = LBound(Criteria)
    m = UBound(Criteria)
    If m + 1 - L > 0 Then                                                                   ' Ensure some criteria
        If SelectDir(HdlWnd, Result, "", "Select Directory") Then                           '  Put all filenames
            Result = Trim$(Result)
            If Right$(Result, 1) <> "\" Then
                Result = Result & "\"
            End If
            Do Until L > m
                FileMatch = Trim$(Dir$(Result & Criteria(L)))                         ' into in array
                Do Until (FileMatch = "")
                    If FirstMatch Then
                        ReDim FileArray(1 To 1) As String                                   ' initial decl of array
                        FirstMatch = False
                    Else
                        ReDim Preserve FileArray(LBound(FileArray) To UBound(FileArray) + 1)
                    End If
                    FileArray(UBound(FileArray)) = Result & FileMatch
                    FileMatch = Trim$(Dir$)
                Loop
                L = L + 1
            Loop
            If FirstMatch Then
                MsgBox "There are no matching files in this directory.", 48, "Message"
            Else
                i = LBound(FileArray)
                ArrayMax = UBound(FileArray)
                Do Until i = ArrayMax                                                       ' insertion sort
                    j = i + 1                                                               ' of the array
                    k = i
                    Do Until j > ArrayMax
                        If FileArray(j) < FileArray(k) Then
                            k = j
                        End If
                        j = j + 1
                    Loop
                    If k <> i Then
                        temp = FileArray(k)
                        FileArray(k) = FileArray(i)
                        FileArray(i) = temp
                    End If
                    i = i + 1
                Loop                                                                        ' end of sort
                Completed = True
            End If
        End If
    End If
Exit_MatchFiles:
    MatchFiles = Completed
    Exit Function
Err_MatchFiles:
    MsgBox Error$
    Resume Exit_MatchFiles
End Function

Function SelectDir(ByVal hWnd As Integer, stControl As String, ByVal StartDir As String, ByVal Title As String) As Variant
'---------------------------------------------------------------------------------
'
' This function opens a window whereby a directory can be selected.
' Parameters:
'       hWND        : handle for current window - should be determined by form property Me.hWnd
'       StartDir    : Any directory desired for the start
'
' Returns
'       stControl   : the selected directory
'       selectDir   : an integer indicating an error or not
'

    On Error GoTo SelectDir_Err
    Dim stDir As String, iTmp As Integer, ReturnValue  As Variant
    ReturnValue = False
    stDir = Trim$(ConvertNull(StartDir, ""))
    stDir = stDir & Chr$(0) & Space$(255)
    'If Get_Directory(Hwnd, Title, stDir) <> 0 Then
    '    stDir = Left(stDir, InStr(stDir, Chr$(0)) - 1)
    '    stControl = stDir
    '    ReturnValue = True
    'End If
Exit_SelectDir:
    SelectDir = ReturnValue
    Exit Function
SelectDir_Err:
    MsgBox Error$
    ReturnValue = False
    Resume Exit_SelectDir
End Function

Function StringFromSz(szTmp As String) As String
    
                                                        
    Dim ich As Integer                                  '  If string terminates with nulls, return a truncated string.
    ich = InStr(szTmp, Chr$(0))
    If ich Then
        StringFromSz = Left$(szTmp, ich - 1)
    Else
        StringFromSz = szTmp
    End If
End Function