Option Compare Database
Option Explicit

Global HourGlassCount As Integer

' Prevent resizing while already resizing
Global m_blResize As Boolean


'*****************************************************************************************************************************
'Purpose:       This routine 'undoes' the effect of using the wheel mouse.  It simply goes back to the previous record which
'                should be the only record.
'Parameters:    None
'Returns:       None
'Created By:    Andrew Rogers
'Created On:    Wed 2/Oct/2002
'Comments:      None
'*****************************************************************************************************************************
Public Sub DontUseWheelMouse(frm As Form)
On Error GoTo DontUseWheelMouse_Err
 
 If frm.NewRecord Then SendKeys ("{PGUP}")
  

DontUseWheelMouse_Exit:
  On Error Resume Next
  Exit Sub

DontUseWheelMouse_Err:
  Call DisplayErrMsg("DontUseWheelMouse")
  Resume DontUseWheelMouse_Exit

End Sub

Public Sub ShowHourGlass()
On Error Resume Next
  HourGlassCount = HourGlassCount + 1
  DoCmd.Hourglass True
  'Debug.Print HourGlassCount
End Sub

Public Sub HideHourGlass(Optional Force As Boolean = False)
  On Error Resume Next
  If Force Then HourGlassCount = 0
  HourGlassCount = HourGlassCount - 1
  'Debug.Print HourGlassCount
  If HourGlassCount <= 0 Then
    DoCmd.Hourglass False
    HourGlassCount = 0
  End If
End Sub

Public Function ShowVersionInformation(Optional ModuleName As String)
  
  DoCmd.OpenForm "Version Information", , , , , , ModuleName
  
End Function

' Open Help File by using F1 Help key
Public Function ShowHelp(Optional uContext As Long)
    SendKeys "{F1}"
End Function

' Sets (and creates if necessary) application properties like Icons and Titles
Public Function AddAppProperty(strName As String, varType As Variant, varValue As Variant) As Integer
  
  Dim prp As Property
  Const conPropNotFoundError = 3270

  On Error GoTo AddProp_Err
  CurrentDb.Properties(strName) = varValue

AddAppProperty = True

AddProp_Bye:
  Exit Function

AddProp_Err:
  If Err = conPropNotFoundError Then
    Set prp = CurrentDb.CreateProperty(strName, varType, varValue)
    CurrentDb.Properties.Append prp
    Resume
  Else
    AddAppProperty = False
    Resume AddProp_Bye
  End If
End Function


Public Sub SetPropertiesForAllForms(Optional MenuBar, Optional ToolBar, Optional ShortcutMenuBar, Optional HelpFile, Optional HelpTopic, Optional Override = False)

  Dim dbs As Database, ctr As Container, doc As Document, f As Form

  Set dbs = CurrentDb
  Set ctr = dbs.Containers!Forms
  For Each doc In ctr.Documents
    
    DoCmd.OpenForm doc.Name, acDesign
    Set f = Forms(doc.Name)
    
    If Not IsMissing(MenuBar) And (Override Or f.MenuBar = "") Then f.MenuBar = MenuBar
    If Not IsMissing(ToolBar) And (Override Or f.ToolBar = "") Then f.ToolBar = ToolBar
    If Not IsMissing(ShortcutMenuBar) And (Override Or f.ShortcutMenuBar = "") Then f.ShortcutMenuBar = ShortcutMenuBar
    If Not IsMissing(HelpFile) And (Override Or f.HelpFile = "") Then f.HelpFile = HelpFile
    If Not IsMissing(HelpTopic) And (Override Or f.HelpContextId = 0) Then f.HelpContextId = HelpTopic
    
    'If F.MenuBar = "" Then F.MenuBar = MenuBar
    'If F.ToolBar = "" Then F.ToolBar = ToolBar
    'If F.ShortcutMenuBar = "" Then F.ShortcutMenuBar = ShortcutMenuBar
  
    DoCmd.Save acForm, doc.Name
    DoCmd.Close acForm, doc.Name
    
  Next doc
  
  Set dbs = Nothing

  
End Sub


Public Sub SetPropertiesForAllReports(Optional MenuBar, Optional ToolBar, Optional ShortcutMenuBar, Optional HelpFile, Optional HelpTopic, Optional Override = False)

  Dim dbs As Database, ctr As Container, doc As Document, r As Report

  Set dbs = CurrentDb
  Set ctr = dbs.Containers!Reports
  For Each doc In ctr.Documents
    
    DoCmd.OpenReport doc.Name, acDesign
    Set r = Reports(doc.Name)
    
    If Not IsMissing(MenuBar) And (Override Or r.MenuBar = "") Then r.MenuBar = MenuBar
    If Not IsMissing(ToolBar) And (Override Or r.ToolBar = "") Then r.ToolBar = ToolBar
    If Not IsMissing(ShortcutMenuBar) And (Override Or r.ShortcutMenuBar = "") Then r.ShortcutMenuBar = ShortcutMenuBar
    If Not IsMissing(HelpFile) And (Override Or r.HelpFile = "") Then r.HelpFile = HelpFile
    If Not IsMissing(HelpTopic) And (Override Or r.HelpContextId = 0) Then r.HelpContextId = HelpTopic
    
    DoCmd.Save acReport, doc.Name
    DoCmd.Close acReport, doc.Name
    
  Next doc
  
  Set dbs = Nothing
  
End Sub
Public Sub ShowPropertiesForAllForms()

  Dim dbs As Database, ctr As Container, doc As Document, f As Form

  Set dbs = CurrentDb
  Set ctr = dbs.Containers!Forms
  For Each doc In ctr.Documents
    
    DoCmd.OpenForm doc.Name, acDesign
    Set f = Forms(doc.Name)
    
    Debug.Print f.Name & ": Menu: " & f.MenuBar & ": Toolbar: " & f.ToolBar & ": Shortcut: " & f.ShortcutMenuBar
    Debug.Print "  Help: " & f.HelpFile & ": Context: " & f.HelpContextId
  
    DoCmd.Close acForm, doc.Name
    
  Next doc
  
  Set dbs = Nothing

  
End Sub

Public Function MsgBox2(Prompt, Optional Buttons, Optional Title) As Long
On Error GoTo MsgBox2_Err

  ReturnVar = vbNo
  
  If IsMissing(Buttons) Then
    Buttons = vbOKOnly
  End If
  
  If IsMissing(Title) Then
    Title = ""
  End If
  
  DoCmd.OpenForm "MsgboxForm", , , , , acDialog, Buttons & "|" & Title & "|" & Prompt
  
  MsgBox2 = ReturnVar

MsgBox2_Exit:
  Exit Function
  
MsgBox2_Err:
  Call DisplayErrMsg("MsgBox2")
  Resume MsgBox2_Exit
  
End Function

' Automatically Close and reopen the active report or form in design mode
' Used to design the active object when under source control

Public Function DesignActiveObject()
  On Error GoTo DesignActiveForm_Err

  Dim intCurrentType As Integer
  Dim strCurrentName As String

  intCurrentType = Application.CurrentObjectType
  strCurrentName = Application.CurrentObjectName
  
  DoCmd.Close intCurrentType, strCurrentName
  
  Select Case intCurrentType
  
    Case acTable
      DoCmd.OpenTable strCurrentName, acViewDesign
      
    Case acQuery
      DoCmd.OpenQuery strCurrentName, acViewDesign
      
    Case acForm
      DoCmd.OpenForm strCurrentName, acDesign
      
    Case acReport
      DoCmd.OpenReport strCurrentName, acViewDesign
      
    Case acMacro
    
    Case acModule
      DoCmd.OpenModule strCurrentName
    
  End Select
  
  GoTo DesignActiveForm_Exit
  
  On Error Resume Next
  
  Dim Ob As String
  Ob = Screen.ActiveForm.Name
  
  If Err.Number = 0 Then
    DoCmd.Close acForm, Ob
    DoCmd.OpenForm Ob, acDesign
  Else
    On Error Resume Next
    Ob = Screen.ActiveReport.Name
    If Err.Number = 0 Then
      DoCmd.Close acReport, Ob
      DoCmd.OpenReport Ob, acViewDesign
    Else
      MsgBox "Confused: " & Err.Description, vbExclamation
    End If
  End If
  
DesignActiveForm_Exit:
  Exit Function
  
DesignActiveForm_Err:
  Call DisplayErrMsg("DesignActiveForm")
  
End Function


'***************************************************************************
'Purpose:       Adds error checking code to the current procedure
'Parameters:    None
'Returns:       None
'               Zero-length string if clipboard is empty or not text
'Created By:    Andrew Rogers
'Created On:
'Comments:      Requires that the procedure name to which the the error
'               checking code is to be added is SELECTED.
'***************************************************************************
Public Function AddErrorCheckingCode()
  Dim ProcedureName  As String, ProcType As String
  Dim Mdl As Module
  Dim StartLine As Long, LastLine As Long
  
  SendKeys "^c", True
  
  Set Mdl = Modules(0)
  
  'ProcedureName = InputBox("Enter the procedure name.")
  'If ProcedureName = "" Then Exit Function
  
  ProcedureName = GetClipboardText
  
  'ProcedureName = "Test"
  
  If ProcedureName = "" Then
    MsgBox "Select the procedure name first.", vbInformation
    Exit Function
  End If
  
  Response = MsgBox("Add header for: " & ProcedureName & CRLF(2) & "YES: Include header" & CRLF(1) & "NO: Exclude header", vbYesNoCancel + vbDefaultButton2 + vbQuestion)
  
  If Response = vbCancel Then Exit Function
  
  StartLine = Mdl.ProcBodyLine(ProcedureName, vbext_pk_Proc)
  ProcType = StringParse(Mdl.Lines(StartLine, 1), 2, " ")
  
  If Response = vbYes Then
    Q = ""
    Q = Q & "'*****************************************************************************************************************************" & CRLF(1)
    Q = Q & "'Purpose:       -" & CRLF(1)
    Q = Q & "'Parameters:    None" & CRLF(1)
    Q = Q & "'Returns:       None" & CRLF(1)
    Q = Q & "'Created By:    Andrew Rogers" & CRLF(1)
    Q = Q & "'Created On:    " & Format(Now, "ddd d/mmm/yyyy") & CRLF(1)
    Q = Q & "'Comments:      None" & CRLF(1)
    Q = Q & "'*****************************************************************************************************************************"
    
    Mdl.InsertLines StartLine, Q
  End If
  
  StartLine = Mdl.ProcBodyLine(ProcedureName, vbext_pk_Proc)
  
  Q = ""
  Q = Q & "  On Error Goto " & ProcedureName & "_Err"
  Mdl.InsertLines StartLine + 1, Q
  
  LastLine = Mdl.ProcStartLine(ProcedureName, vbext_pk_Proc) + Mdl.ProcCountLines(ProcedureName, vbext_pk_Proc)
  
  Q = ""
  Q = Q & ProcedureName & "_Exit:" & CRLF(1)
  Q = Q & vbTab & "On Error Resume Next" & CRLF(1)
  Q = Q & vbTab & "Exit " & ProcType & CRLF(2)
  Q = Q & ProcedureName & "_Err:" & CRLF(1)
  Q = Q & vbTab & "Call DisplayErrMsg(""" & ProcedureName & """)" & CRLF(1)
  Q = Q & vbTab & "Resume " & ProcedureName & "_Exit" & CRLF(1)
  
  Mdl.InsertLines LastLine - 1, Q
  
End Function



Public Sub FixSccStatus()

  Dim dbs As Database, ctr As Container, doc As Document
  Dim i As Integer
  
  Dim intCurrentType As Integer
  Dim strCurrentName As String

  intCurrentType = Application.CurrentObjectType
  strCurrentName = Application.CurrentObjectName
  
  Response = MsgBox("Push Yes to fix all.  Push No to fix selected.", vbYesNoCancel + vbExclamation)
  
  Set dbs = CurrentDb
  'Stop
  
  If Response = vbYes Then
      
    For intCurrentType = 1 To 5
      Select Case intCurrentType
      
        Case acForm
          Set ctr = dbs.Containers!Forms
        Case acMacro
          Set ctr = dbs.Containers!scripts
        Case acReport
          Set ctr = dbs.Containers!Reports
        Case acModule
          Set ctr = dbs.Containers!Modules
        Case acTable, acQuery
          Set ctr = dbs.Containers!Tables
        
      End Select
      
      'Stop
      For Each doc In ctr.Documents
        Call FixSccObjectStatus(intCurrentType, doc.Name, True)
      
      Next
    Next
  
  Else
    Call FixSccObjectStatus(intCurrentType, strCurrentName, False)
  End If
  
End Sub


Private Sub FixSccObjectStatus(oType, oName, Quiet As Boolean)
  On Error GoTo FixSccObjectStatus_Err

  SysCmd acSysCmdSetStatus, "Checking: " & oName
  Dim dbs As Database, ctr As Container, doc As Document
  
  Set dbs = CurrentDb
  'Stop
  
  Select Case oType
  
    Case acForm
      Set ctr = dbs.Containers!Forms
    Case acMacro
      Set ctr = dbs.Containers!scripts
    Case acReport
      Set ctr = dbs.Containers!Reports
    Case acModule
      Set ctr = dbs.Containers!Modules
    Case acTable, acQuery
      Set ctr = dbs.Containers!Tables
    
  End Select
  
  Set doc = ctr.Documents(oName)
  If doc.Properties("SccStatus").Value <> 1 Then
    If Not Quiet Then
      Response = MsgBox("SccStatus = " & doc.Properties("SccStatus").Value & ".  Set to 1", vbExclamation + vbYesNo)
    Else
      Response = vbYes
    End If
    If vbYes Then doc.Properties("SccStatus").Value = 1
  Else
    If Not Quiet Then MsgBox "SccStatus looks OK.", vbInformation
  End If
  
FixSccObjectStatus_Exit:
  On Error Resume Next
  Exit Sub

FixSccObjectStatus_Err:
  If Err.Number <> 3270 Then Call DisplayErrMsg("FixSccObjectStatus")
  Resume FixSccObjectStatus_Exit


End Sub


Public Sub CodeTiming(Optional CodeDescription As String = "-> ", Optional start As Boolean = False)

  Static LastTime
  
  If start Then
    LastTime = Timer
    Debug.Print "Starting " & CodeDescription & " ..."
  Else
    Debug.Print CodeDescription & " " & Format(Timer - LastTime, "0.00") & " secs"
    LastTime = Timer
  End If
  
End Sub

'*******************
' Show and Hide Development User Interface
'*******************
Public Sub UserMode(HideDev As Boolean)
  On Error Resume Next
  
  ' Can Check if runtime using SysCmd(acSysCmdRuntime)
  ' If SysCmd(acSysCmdRuntime) = False Then
        'The user is using a full blown version of MS Access
  
  If HideDev Then
    DevelopmentModeSet = False
    'select the navigation pange
    Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
    'hide the selected object
    Call DoCmd.RunCommand(acCmdWindowHide)
    
    DoCmd.ShowToolbar "Database", acToolbarNo
    DoCmd.ShowToolbar "Form View", acToolbarNo
    DoCmd.ShowToolbar "Print Preview", acToolbarWhereApprop
  Else
    DevelopmentModeSet = True
    'DoCmd.RunCommand acCmdWindowUnhide
    Call DoCmd.SelectObject(acTable, , True)
    DoCmd.ShowToolbar "Database", acToolbarYes
    DoCmd.ShowToolbar "Form View", acToolbarYes
    DoCmd.ShowToolbar "Print Preview", acToolbarYes
  End If
  
End Sub

Private Function LoadRibbons(strRibbon)

        On Error GoTo Error1

        Dim rs As DAO.Recordset
        Dim strRibbonData As String

      '  Set RS = CurrentDb.OpenRecordset("SELECT * FROM USysRibbons")
        strRibbonData = DLookup("[RibbonXML]", "USysRibbons", "[RibbonName] = """ & strRibbon & """")
        Application.LoadCustomUI strRibbon, strRibbonData

   '  Do Until RS.EOF
'
 '        If RS("RibbonName").Value = strRibbon Then
              ' Ribbon found: Load it and exit
  '          Application.LoadCustomUI strRibbon, RS("RibbonXML").Value
   '         Exit Do
    '    End If

     '    RS.MoveNext

    ' Loop

Error1_Exit:

     On Error Resume Next
    ' RS.Close
     Set rs = Nothing
     Exit Function

Error1:

     Select Case Err
         Case 32609
         ' Ribbon already loaded, do nothing and exit
     Case Else
         MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error", Err.HelpFile, Err.HelpContext
     End Select

     Resume Error1_Exit

 End Function
 

Sub glrMinWindowSize(frm As Form, Optional intTotalFormHeight As Long, _
  Optional intTotalFormWidth As Long, Optional bNoWidth As Boolean = False)
    ' copied originally from Access Help files
    ' modified by Charlotte Foust
    ' From http://www.utteraccess.com/forum/lofiversion/index.php/t2041167.html
    
    Dim intWindowHeight As Long
    Dim intWindowWidth As Long
 
    ' Stop reprocess during resizing
    m_blResize = True
    
    ' Determine form's height.
    If intTotalFormHeight = 0 Then
        intTotalFormHeight = frmHeight(frm)
    End If
    
    ' Determine form's width.
    If intTotalFormWidth = 0 Then
        intTotalFormWidth = frm.Width
    End If
    
    ' Determine window's height and width.
    intWindowHeight = frm.InsideHeight
    intWindowWidth = frm.InsideWidth
    
    If (bNoWidth And (intWindowWidth <> intTotalFormWidth)) Or _
           intWindowWidth < intTotalFormWidth Then
        frm.InsideWidth = intTotalFormWidth
    End If
    
    If intWindowHeight < intTotalFormHeight Then
        frm.InsideHeight = intTotalFormHeight
    End If
    
    ' Allow resizing
    m_blResize = False
End Sub

Function IsSectionThere(frm As Form, s As Long) As Boolean
    Dim nm As String
    On Error GoTo EH
    nm = frm.Section(s).Name  '*** This will throw an error if the section is not there
    IsSectionThere = True
    Exit Function
EH:
    If Err.Number = 2462 Then     '*** handle specific error for section not found "seamlessly"
        IsSectionThere = False    '*** You can specify the exact error number to trap, if you want...
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description ' Display message for all others
    End If
End Function

Function frmHeight(frm As Form) As Long
    On Error Resume Next
    Dim intTotalFormHeight As Long
    
    Dim intHeightHeader As Long
    Dim intHeightDetail As Long
    Dim intHeightFooter As Long

    ' Calculate
    ' Just use Resume next instead of testing sections
    'If IsSectionThere(frm, acHeader) Then intHeightHeader = frm.Section(acHeader).Height
    intHeightHeader = frm.Section(acHeader).Height
    intHeightDetail = frm.Section(acDetail).Height
    'If IsSectionThere(frm, acFooter) Then intHeightFooter = frm.Section(acFooter).Height
    intHeightFooter = frm.Section(acFooter).Height
    intTotalFormHeight = intHeightHeader + intHeightDetail + intHeightFooter
    
    frmHeight = intTotalFormHeight
End Function