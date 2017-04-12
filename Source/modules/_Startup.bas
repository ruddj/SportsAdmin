Option Compare Database
Option Explicit

Public Function Startup()
On Error GoTo Startup_Err

    Dim FilePath As String
    
  Application.MenuBar = "Sports Menu"
  Call UserMode(True)
  
  Call InitialiseWaitMessage
  DoCmd.RunMacro "ShowPleaseWait"
  
  ' If running in runtime mode add current location to Trusted Location if not already
  If SysCmd(acSysCmdRuntime) Then
    FilePath = Application.CurrentProject.Path
    Call AddTrustedLocation(FilePath, "Sports Admin Application", True)
  End If
  
 'DoCmd.RunCommand acCmdWindowHide
 ' DoCmd.ShowToolbar "Database", acToolbarNo
 ' DoCmd.ShowToolbar "Form View", acToolbarNo
 ' DoCmd.ShowToolbar "Print Preview", acToolbarWhereApprop
  Call UserMode(True)

  Call CheckInventoryAttached
    
  'CurrentDb.Properties("AppTitle") = "Sports Administrator v" & VersionNumber
  Application.RefreshTitleBar

  Call UpdateEventCompetitorAge
  
  DoCmd.RunMacro "ClosePleaseWait"
  
  DoCmd.OpenForm "Main Menu"
  
  ' Create Report Right Click menu
  CreateReportShortcutMenu

Startup_Exit:
  Exit Function
  
Startup_Err:
  MsgBox "An error has occurred in [Startup]: " & Err.Description, vbCritical
  Resume Startup_Exit
  
End Function