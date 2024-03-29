Option Compare Database
Option Explicit

Const REPORT_MENU As String = "cmdReportRightClick"

Public Function CreateReportList(Optional CalledFromForm As Boolean)
On Error GoTo CreateReportList_err

  Dim NumberReports As Integer, X As Integer, ObjectType As String, ReportName As String
  Dim db As Database, Rs As Recordset
  
  ObjectType = Application.CurrentObjectType
  ReportName = Application.CurrentObjectName
  
  'MsgBox (ObjectType)
  
  Set db = CurrentDb
  Set Rs = db.OpenRecordset("ReportList", dbOpenDynaset)
  
  DoCmd.SetWarnings False
  DoCmd.RunSQL "UPDATE ReportList SET ReportList.Open = false"
  
  NumberReports = Reports.Count   ' Count number of reports.
  ' If no reports open or last report is being closed then
  If NumberReports = 0 Or (NumberReports = 1 And ObjectType = acReport And Not CalledFromForm) Then
    DoCmd.Close acForm, "ReportsPopup"
  Else
    For X = 0 To NumberReports - 1
'      DoCmd.SelectObject acReport, Reports(x).Name, False
'      DoCmd.RunCommand acCmdPreviewTwoPages
      Rs.FindFirst "[ReportName]=""" & Reports(X).Name & """"
      If Rs.NoMatch Then
        Rs.AddNew
        Rs![ReportName] = Reports(X).Name
        If VarEmpty(Reports(X).Caption) Then
          Rs![ReportCaption] = Reports(X).Name
        Else
          Rs![ReportCaption] = Reports(X).Caption
        End If
      Else
        Rs.Edit
      End If
      If Rs!ReportName = ReportName Then ' Pass ReportName to the procedure when a report is closed
        Rs!Open = False
      Else
        Rs!Open = True
      End If
      
      Rs.Update
    Next X
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

Public Function CreateReportShortcutMenu()
    Dim cmbRightClick As Office.CommandBar
    Dim cmbControl As Office.CommandBarControl
 
    Call DeleteReportMenu

   ' Create the shortcut menu.
    Set cmbRightClick = CommandBars.Add(Name:=REPORT_MENU, Position:=msoBarPopup, MenuBar:=False, Temporary:=True)
 
    With cmbRightClick
         
        ' Add the Print command.
        Set cmbControl = .Controls.Add(msoControlButton, 2521, , , True)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Quick Print"
         
        ' Add the Print command.
        Set cmbControl = .Controls.Add(msoControlButton, 15948, , , True)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Select Pages"
         
        ' Add the Page Setup... command.
        Set cmbControl = .Controls.Add(msoControlButton, 247, , , True)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Page Setup"
         
        ' Add the Mail Recipient (as Attachment)... command.
        Set cmbControl = .Controls.Add(msoControlButton, 2188, , , True)
        ' Start a new group.
        cmbControl.BeginGroup = True
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Email Report as an Attachment"
         
        ' Add the PDF or XPS command.
        Set cmbControl = .Controls.Add(msoControlButton, 12499, , , True)
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Save as PDF/XPS"
         
        ' Add the Close command.
        Set cmbControl = .Controls.Add(msoControlButton, 923, , , True)
        ' Start a new group.
        cmbControl.BeginGroup = True
        ' Change the caption displayed for the control.
        cmbControl.Caption = "Close Report"
    End With
     
    Set cmbControl = Nothing
    Set cmbRightClick = Nothing
End Function


Sub DeleteReportMenu()

    On Error Resume Next
    Application.CommandBars(REPORT_MENU).Delete
    
End Sub