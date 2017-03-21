Option Compare Database

'Command to pass HTMLHelp()
Public Const HH_DISPLAY_TOPIC = &H0 ' Display the help file.
Public Const HH_DISPLAY_TOC = &H1 ' Display the table of contents.
Public Const HH_DISPLAY_INDEX = &H2 ' Display the index.
Public Const HH_DISPLAY_SEARCH = &H3 ' Display full text search.
Public Const HH_HELP_CONTEXT = &HF ' Display mapped numeric value in dwData.
Public Const HH_CLOSE_ALL = &H12 ' Close the help file.

Public Const HH_SET_WIN_TYPE As Long = &H4
Public Const HH_GET_WIN_TYPE As Long = &H5
Public Const HH_GET_WIN_HANDLE As Long = &H6
Public Const HH_DISPLAY_TEXT_POPUP As Long = &HE
Public Const HH_TP_HELP_CONTEXTMENU As Long = &H10
Public Const HH_TP_HELP_WM_HELP As Long = &H11

Global strCHMHelp As String

Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
(ByVal hwndCaller As Long, ByVal pszFile As String, _
ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Function ShowHelp(uContext As Long)
    Call HtmlHelp(0, strCHMHelp, HH_HELP_CONTEXT, uContext)
End Function

Public Sub SetHelp()
    Dim FilePath As String
    FilePath = Application.CurrentProject.Path
    strCHMHelp = FilePath & "\SportsAdmin.chm"
End Sub

Public Function CloseHelp()
    Call HtmlHelp(0, strCHMHelp, HH_CLOSE_ALL, 0)
End Function