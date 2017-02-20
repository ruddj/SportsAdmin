Option Compare Database
Option Explicit


Public Function CreateReportList(Optional CalledFromForm As Boolean)
On Error GoTo CreateReportList_err

  Dim NumberReports As Integer, x As Integer, ObjectType As String, ReportName As String
  Dim db As Database, rs As Recordset
  
  ObjectType = Application.CurrentObjectType
  ReportName = Application.CurrentObjectName
  
  'MsgBox (ObjectType)
  
  Set db = CurrentDb
  Set rs = db.OpenRecordset("ReportList", dbOpenDynaset)
  
  DoCmd.SetWarnings False
  DoCmd.RunSQL "UPDATE ReportList SET ReportList.Open = false"
  
  NumberReports = Reports.Count   ' Count number of reports.
  ' If no reports open or last report is being closed then
  If NumberReports = 0 Or (NumberReports = 1 And ObjectType = acReport And Not CalledFromForm) Then
    DoCmd.Close acForm, "ReportsPopup"
  Else
    For x = 0 To NumberReports - 1
'      DoCmd.SelectObject acReport, Reports(x).Name, False
'      DoCmd.RunCommand acCmdPreviewTwoPages
      rs.FindFirst "[ReportName]=""" & Reports(x).Name & """"
      If rs.NoMatch Then
        rs.AddNew
        rs![ReportName] = Reports(x).Name
        If VarEmpty(Reports(x).Caption) Then
          rs![ReportCaption] = Reports(x).Name
        Else
          rs![ReportCaption] = Reports(x).Caption
        End If
      Else
        rs.Edit
      End If
      If rs!ReportName = ReportName Then ' Pass ReportName to the procedure when a report is closed
        rs!Open = False
      Else
        rs!Open = True
      End If
      
      rs.Update
    Next x
    Forms![ReportsPopUp]!ReportList.Requery
  End If
  
CreateReportList_Exit:
  DoCmd.SetWarnings True
  Exit Function
  
CreateReportList_err:
  If Err.Number = 2450 Then ' cant find ReportsPopUp because it has been closed .  that ok just continue
    ' do nothing
  Else
    MsgBox ("An error has occured in [CreateReportList_err]: " & Err.Description)
  End If
  
  GoTo CreateReportList_Exit
  
End Function