Option Compare Database
Option Explicit

Public Function Startup()
On Error GoTo Startup_Err

  Application.MenuBar = "Sports Menu"
  'DoCmd.RunCommand acCmdWindowHide
  Call InitialiseWaitMessage
  DoCmd.RunMacro "ShowPleaseWait"
  
  'DoCmd.ShowToolbar "Database", acToolbarNo
  'DoCmd.ShowToolbar "Form View", acToolbarNo
  'DoCmd.ShowToolbar "Print Preview", acToolbarWhereApprop
  
  Call CheckInventoryAttached
    
  'CurrentDb.Properties("AppTitle") = "Sports Administrator v" & VersionNumber
  Application.RefreshTitleBar

  Call UpdateEventCompetitorAge
  
  DoCmd.RunMacro "ClosePleaseWait"
  
  DoCmd.OpenForm "Main Menu"

Startup_Exit:
  Exit Function
  
Startup_Err:
  MsgBox "An error has occurred in [Startup]: " & Err.description, vbCritical
  Resume Startup_Exit
  
End Function