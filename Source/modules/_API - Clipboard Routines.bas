Option Compare Database
Option Explicit

'Windows API declarations
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hWnd As LongPtr, ByVal lpClassName As String, _
     ByVal nMaxCount As LongPtr) As Long

Private Declare PtrSafe Function GetWindow Lib "user32" _
    (ByVal hWnd As LongPtr, ByVal wCmd As Long) As Long

Private Declare PtrSafe Function OpenClipboard Lib "user32" _
    (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As LongPtr

Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr

Private Declare PtrSafe Function GlobalLock Lib "kernel32" _
    (ByVal hMem As LongPtr) As LongPtr

Private Declare PtrSafe Function lstrcpy Lib "kernel32" _
    (ByVal lpString1 As Any, _
     ByVal lpString2 As Any) As LongPtr

Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long

Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

'Constants used by Windows API calls
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
 
Function GetClipboardText() As String
'***************************************************************************
'Purpose:       To fetch text data from the clipboard
'Parameters:    None
'Returns:       A string containing the clipboard text;
'               Zero-length string if clipboard is empty or not text
'Created By:    Rob Smith
'Created On:    20 Nov 95
'Comments:      If only we had the VB function Clipboard.GetText() !
'***************************************************************************
 
Dim lngMemBlockHandle As LongPtr
Dim lngMemPointer As LongPtr
Dim strText As String
Dim lngRetVal As LongPtr


' Open the Clipboard
If OpenClipboard(0&) = 0 Then
   MsgBox "Could not open the Clipboard."
   Exit Function
End If

' Obtain the handle to the global memory block that references the text
lngMemBlockHandle = GetClipboardData(CF_TEXT)
If lngMemBlockHandle = 0 Then
   MsgBox "Could not allocate memory."
   GoTo GetClipboardText_Exit
End If

' Lock Clipboard memory so we can reference the data string
lngMemPointer = GlobalLock(lngMemBlockHandle)

If lngMemPointer <> 0 Then
   strText = Space$(MAXSIZE)
   
   'Copy data from lngMemPointer into strText
   lngRetVal = lstrcpy(strText, lngMemPointer)
   
   'Unlock Clipboard memory block
   lngRetVal = GlobalUnlock(lngMemBlockHandle)

   'Peel off the Null termination character
   strText = Mid(strText, 1, InStr(1, strText, Chr$(0), 0) - 1)
Else
   MsgBox "Could not lock memory to copy string from."
End If
 
GetClipboardText_Exit:
    lngRetVal = CloseClipboard()
    GetClipboardText = strText

End Function