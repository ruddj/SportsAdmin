Option Compare Database
Option Explicit

Public Sub LimitedLanes_NoData()

  Dim msg As String, Response As Integer
  msg = "No competitors were found in the 'limited lane' events you specified." & vbLf & vbCr & vbLf & vbCr
  msg = msg & "CHECK:" & vbLf & vbCr
  msg = msg & "1. That you have selected (ticked) the correct events." & vbLf & vbCr
  msg = msg & "2. That you have selected the correct criteria. Normally lists are generated for 'Active' events." & vbLf & vbCr
  msg = msg & "3. That competitors have been added to the event(s)." & vbLf & vbCr
  msg = msg & "4. That the competitors have been allocated a valid lane.  Any competitors in lane 0 will not be displayed in the marshalling lists." & vbLf & vbCr
  
  
  Response = MsgBox(msg, vbInformation)

End Sub

Public Sub UnLimitedLanes_NoData()

  Dim msg As String, Response As Integer
  msg = "No competitors were found in the 'unlimited lane / field' events you specified." & vbLf & vbCr & vbLf & vbCr
  msg = msg & "CHECK:" & vbLf & vbCr
  msg = msg & "1. That you have selected (ticked) the correct events." & vbLf & vbCr
  msg = msg & "2. That you have selected the correct criteria. Normally lists are generated for 'Active' events." & vbLf & vbCr
  msg = msg & "3. That competitors have been added to the event(s)." & vbLf & vbCr
  
  Response = MsgBox(msg, vbInformation)

End Sub


Function PrintOpenReports()

On Error GoTo PrintOpen_Click_Err

    Dim X As Integer, NumberReports As Variant

    NumberReports = Reports.Count   ' Count number of reports.

    For X = 0 To NumberReports - 1
        
        DoCmd.OpenReport Reports(X).Name, A_NORMAL

    Next X

PrintOpen_Click_Exit:
    Exit Function

PrintOpen_Click_Err:
    MsgBox ("Error in PrintOpen_Click: " & Error$)
    GoTo PrintOpen_Click_Exit


End Function


Public Sub PreviewReport(ReportName As String, Optional PreviewOption)
On Error GoTo PreviewReport_Err

  DoCmd.OpenReport ReportName, A_PREVIEW
  DoCmd.Maximize
  If Not IsMissing(PreviewOption) Then DoCmd.RunCommand PreviewOption

PreviewReport_Exit:
  Exit Sub
  
PreviewReport_Err:
  MsgBox "An error has occurred in [PreviewReport]: " & Err.Description, vbCritical
  Resume PreviewReport_Exit
  
End Sub