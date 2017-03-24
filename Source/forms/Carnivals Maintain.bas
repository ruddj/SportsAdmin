Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =10658
    ItemSuffix =26
    Left =-20805
    Top =2715
    Right =-10140
    Bottom =10065
    HelpContextId =30
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0x2e9b3a042dc7e140
    End
    RecordSource ="MiscellaneousLocal"
    Caption ="Maintain Carnivals"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin Section
            Height =7344
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    ColumnCount =4
                    Left =144
                    Top =283
                    Width =8618
                    Height =6126
                    BorderColor =12632256
                    Name ="List"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Carnivals.Carnival, Carnivals.[Relative Directory] AS Directo"
                        "ry, Carnivals.Filename, Carnivals.Available FROM Carnivals ORDER BY Carnivals.Av"
                        "ailable DESC , Carnivals.Carnival;"
                    ColumnWidths ="2268;3536;1851;567"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Double click a carnival to work on it."
                    HorizontalAnchor =2
                    VerticalAnchor =2

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =135
                    Top =60
                    Width =1770
                    Height =210
                    FontWeight =700
                    Name ="Text2"
                    Caption ="Available Carnivals"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =9299
                    Top =5970
                    Width =1134
                    Height =510
                    FontWeight =700
                    TabIndex =1
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Return to the previous form."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =9299
                    Top =144
                    Width =1134
                    Height =420
                    TabIndex =2
                    ForeColor =32768
                    Name ="New"
                    Caption ="&New"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Create a new empty carnival."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =79
                    Left =9299
                    Top =623
                    Width =1134
                    Height =420
                    TabIndex =3
                    ForeColor =32768
                    Name ="Copy"
                    Caption ="C&opy"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Copy an exisiting carnival.  Use this if you have already set up a carnival that"
                        " is similar or identical to the new carnival you will be conducting."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =9299
                    Top =3373
                    Width =1134
                    Height =420
                    TabIndex =4
                    ForeColor =128
                    Name ="Delete"
                    Caption ="D&elete"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Delete the selected carnival."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =82
                    Left =9299
                    Top =2154
                    Width =1134
                    Height =420
                    TabIndex =5
                    ForeColor =8404992
                    Name ="Rename"
                    Caption ="&Rename"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Rename the carnival."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =144
                    Top =6861
                    Width =6779
                    Height =239
                    TabIndex =6
                    BackColor =12632256
                    Name ="ActiveCarnival"
                    ControlSource ="ActiveCarnival"
                    FontName ="Arial"
                    ControlTipText ="This shows the carnival that is currently being worked on."
                    VerticalAnchor =1

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    Left =149
                    Top =6633
                    Width =1410
                    Height =225
                    Name ="Text11"
                    Caption ="Active Carnival"
                    FontName ="Arial"
                    VerticalAnchor =1
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =93
                    Left =1226
                    Top =2380
                    Width =1410
                    Height =225
                    Name ="Text13"
                    Caption ="Global Data"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =9299
                    Top =1105
                    Width =1134
                    Height =420
                    TabIndex =7
                    ForeColor =32768
                    Name ="Button16"
                    Caption ="&Add"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Add an exisitng carnival (one you have already created) to the list."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =81
                    Left =9299
                    Top =6615
                    Width =1134
                    Height =510
                    TabIndex =8
                    Name ="Button17"
                    Caption ="&Quit"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Exit the Sports Administrator."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =87
                    Left =9299
                    Top =3978
                    Width =1134
                    Height =960
                    TabIndex =9
                    ForeColor =8404992
                    Name ="MakeActive"
                    Caption ="&Work on selected carnival"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Start working on the carnival you have selected from the list."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =8994
                    Width =0
                    Height =7344
                    Name ="Line19"
                    HorizontalAnchor =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =72
                    Left =9299
                    Top =5074
                    Width =1134
                    Height =510
                    TabIndex =10
                    HelpContextId =30
                    Name ="Help Button"
                    Caption ="&Help"
                    OnClick ="Open Help"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="Help_Button"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9299
                    Top =1700
                    Width =1134
                    Height =420
                    TabIndex =11
                    ForeColor =8404992
                    Name ="ChangeFile"
                    Caption ="Change File"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Change the file associated with the carnival."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9299
                    Top =2607
                    Width =1134
                    Height =420
                    TabIndex =12
                    ForeColor =8404992
                    Name ="CompactCarnivalBut"
                    Caption ="Compact"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Compact the carnival file so that it is not wasting any disk space."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    Left =7058
                    Top =6718
                    Width =1644
                    Height =420
                    TabIndex =13
                    Name ="ImportCarnivalList"
                    Caption ="&Import Carnival List"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="If you have recently upgraded to a new version of the Sports Administrator, use "
                        "the button to import the carnival list from the old version."
                    VerticalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons
Option Explicit

' Form Dimensions
Dim lMinHeight As Long
Dim lMinWidth As Long

Private Sub Button15_Click()
On Error GoTo Err_Button15_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Button15_Click:
    Exit Sub

Err_Button15_Click:
    MsgBox Error$
    Resume Exit_Button15_Click
    
End Sub

Private Sub Button16_Click()

    On Error GoTo Err_Button16_Click
    DoCmd.OpenForm "Carnival Copy", , , , , acDialog, "ADD"
    Me.List.Requery
Exit_Button16_Click:
    Exit Sub
Err_Button16_Click:
    MsgBox Error$
    Resume Exit_Button16_Click
End Sub

Private Sub Button17_Click()
    UserQuit = False
    If MsgBox("Are you sure you want to close the database?", 276, "Warning") = 6 Then
        UserQuit = True
        Call QuitSportsAdministrator(Me)
        
    End If
End Sub

Private Sub ChangeFile_Click()
On Error GoTo ChangeFile_Click_Err
    If IsNull(Me!List) Then
        MsgBox ("You must select a carnival file.")
    Else
        Call locateCarnival
        Me.List.Requery
    End If
    
ChangeFile_Click_Exit:
    Exit Sub
    
ChangeFile_Click_Err:
    MsgBox (Error$)
    GoTo ChangeFile_Click_Exit

End Sub

Private Sub Close_Click()
    On Error GoTo Err_Close_Click
    
    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
  
    'MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

Private Sub Copy_Click()

    On Error Resume Next
    If IsNull(Me.List) Then
        MsgBox "Select a carnival to copy before pressing this button.", vbExclamation, "Message"
    Else
        If DLookup("[Available]", "Carnivals", "[Carnival] = """ & Me.List & """") Then
            DoCmd.OpenForm "Carnival Copy", , , , , acDialog, "COPY"
            Me.List.Requery
        Else
            MsgBox "This carnival is not available so it cannot be copied.", vbExclamation, "Message"
        End If
    End If


End Sub

Private Sub Delete_Click()

    On Error GoTo Err_Delete_Click
    Dim fileName As String, Response As Variant

    If IsNull(Me.List) Then
        MsgBox "Select a carnival to delete before pressing this button.", vbExclamation, "Message"
    Else
        If MsgBox("Are you sure you want to delete the carnival " & Me.List & "?", 276, "Warning") = 6 Then
            fileName = GetCarnivalFullDir(DLookup("[Relative Directory]", "Carnivals", "[Carnival] = """ & Me.List & """"))
            fileName = fileName & DLookup("[Filename]", "Carnivals", "[Carnival] = """ & Me.List & """")
            If FileExists(fileName) Then
                Response = MsgBox("Do you wish to delete the file " & fileName & "?", vbYesNo + vbCritical + vbDefaultButton2, "Delete file?")
                If Response = vbYes Then
                    Kill fileName
                End If
            Else
                MsgBox ("The file associated with this carnival was not in the expected location and so could not be deleted.")
            End If
            DoCmd.SetWarnings False
            DoCmd.RunSQL "DELETE * FROM Carnivals WHERE [Carnival] = """ & Me.List & """;"
            Form_Load
        End If
    End If
Exit_Delete_Click:
    DoCmd.SetWarnings True
    Exit Sub
Err_Delete_Click:
    MsgBox Error$
    Resume Exit_Delete_Click
End Sub



Private Sub Form_Load()

    On Error GoTo Err_Form_Load
    Dim Db As Database, TB As TableDef, fileName As String, RPath As String, RFile As String, FilenamePath As Variant
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "UPDATE DISTINCTROW Carnivals SET Carnivals.Available = FileExists(GetCarnivalFullDir([Relative Directory]) & [Filename]);"
    Me.List.Requery
    Me.List = Null
    Set Db = DBEngine.Workspaces(0).Databases(0)
    Set TB = Db.TableDefs("Competitors")
    fileName = UCase$(Right$(TB.connect, Len(TB.connect) - InStr(TB.connect, "=")))
    FilenamePath = Left$(fileName, Len(fileName) - InStr(ReverseString(fileName), "\") + 1)
    ''RFile = Right$(Filename, Len(Filename) - Len(RPath))
    RPath = GetCarnivalRelDir(FilenamePath)
    RFile = GetCarnivalFile(fileName)
    Me.ActiveCarnival = DLookup("[CArnival]", "Carnivals", "([Filename] = """ & RFile & """) AND ([Relative Directory] = """ & RPath & """) and [Available]")
    UserQuit = False

Exit_Form_Load:
    DoCmd.SetWarnings True
    Exit Sub
Err_Form_Load:
    MsgBox Error$
    Resume Exit_Form_Load
End Sub

Private Sub Form_Open(Cancel As Integer)
    lMinHeight = frmHeight(Me)
    lMinWidth = Me.Width
    
    DoCmd.RunMacro "ClosePleaseWait"

End Sub

Private Sub Form_Resize()
    If Not m_blResize Then Call glrMinWindowSize(Me, lMinHeight, lMinWidth, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If IsNull(Me!ActiveCarnival) And Not UserQuit Then
        MsgBox "You cannot continue until a valid carnival is selected.", vbExclamation, "Message"
        Cancel = True
    End If


End Sub

Private Sub ImportCarnivalList_Click()

On Error GoTo ImportCarnivalList_Click_Err
    
    Dim ReturnVar As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    
    strFilter = ahtAddFilterItem(strFilter, "Carnival Files (*.mde)", "*.mde")
    strFilter = ahtAddFilterItem(strFilter, "Old Files (*.old)", "*.old")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    ReturnVar = ahtCommonFileOpenSave(InitialDir:="", _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate Old Sports.MDE Program file")

    
    If ReturnVar = "" Then
      ' Action Cancelled
    Else
      Dim Db As Database, adb As Database, Crs As Recordset, ocrs As Recordset
      Dim wsp As Workspace, Criteria As String
      
      Set Db = CurrentDb
      Set wsp = DBEngine.Workspaces(0)
      ' Return reference to Another.mdb.
      Set adb = wsp.OpenDatabase(ReturnVar)

      Set Crs = Db.OpenRecordset("Carnivals", dbOpenDynaset)
      Set ocrs = adb.OpenRecordset("Carnivals")
      
      ocrs.MoveFirst
      If ocrs.BOF Then
        MsgBox ("There are no carnivals listed in the old SPORTS application file.")
      Else
        ocrs.MoveFirst
        While Not ocrs.EOF
          Criteria = "[Carnival]=""" & ocrs!Carnival & """"
          Crs.FindFirst Criteria
          If Crs.NoMatch Then
            Crs.AddNew
            Crs!Carnival = ocrs!Carnival
            Crs!fileName = ocrs!fileName
            Crs![Relative Directory] = ocrs![Relative Directory]
            Crs!available = FileExists(GetCarnivalFullDir(ocrs![Relative Directory]) & ocrs![fileName])
            Crs.Update
          End If
          ocrs.MoveNext
        Wend
      
        Me.List.Requery
       End If
    
    End If
   
ImportCarnivalList_Click_Exit:
    Exit Sub
    
ImportCarnivalList_Click_Err:
    MsgBox (Error$)
    GoTo ImportCarnivalList_Click_Exit


End Sub

Private Sub List_DblClick(Cancel As Integer)
'-------------------------------------------------------------------------
'
    On Error GoTo Err_List_DblClick
    'Stop
    Dim fileName As String, Posi  As Variant, HasError As Variant
    DoCmd.Hourglass True
    If IsNull(Me.List) Then
        MsgBox "Select a carnival to connect.", vbExclamation, "Message"
    Else
        PleaseWaitMsg = "Retrieving Carnival Details ..."
        DoCmd.RunMacro "ShowPleaseWait"
        
        If DLookup("[Available]", "Carnivals", "[Carnival] = """ & Me.List & """") Then
            fileName = CarnivalDir(DLookup("[Relative Directory]", "Carnivals", "[Carnival] = """ & Me.List & """")) & DLookup("[Filename]", "Carnivals", "[Carnival] = """ & Me.List & """")
            If Attach_Selected_File2(2, Posi, HasError, fileName) Then
                If Not HasError Then

                    GlobalVariable = SysCmd(acSysCmdSetStatus, "Verifying relationships ... ")
                    Call CheckRelationships(fileName)
                    GlobalVariable = SysCmd(acSysCmdSetStatus, "Finalising carnival selection ... ")
                    Call FinaliseCarnivalSelection

                    Me.ActiveCarnival = DLookup("[CArnival]", "Carnivals", "[Carnival] = """ & Me.List & """ AND [Available]")
                    MsgBox "Selected carnival is now active.", vbInformation
                    GlobalVariable = SysCmd(acSysCmdClearStatus)
                    DoCmd.RunMacro "ClosePleaseWait"
                    DoCmd.Close acForm, "Carnivals Maintain"
                End If
            End If
        Else
            locateCarnival
            Me.List.Requery
        End If
        
    End If
Exit_List_DblClick:
    DoCmd.Hourglass False
    DoCmd.RunMacro "ClosePleaseWait"
    Exit Sub
'xxx
Err_List_DblClick:
    MsgBox Error$
    Resume Exit_List_DblClick

'Err_Creating_Relationships:
'    MsgBox ("Error creating relationship for: Table1=" & R1![First Table] & " Table2=" & R1![Second Table])
'    Resume Next
End Sub

Private Sub locateCarnival()

    On Error GoTo Err_locateCarnival
    Dim MyDb As Database, ITable As Recordset, SpecifiedPath As Variant, TT As TableDef, FTable As Recordset
    Dim DataExists As Variant, MyWS As Workspace, CPath  As Variant, AskUser  As Variant
    Dim Result As Variant, ReturnVal As Variant, Db As Database
    Dim NewDir As String, OldDB As String, NextCarn As String
       
'    Stop
    
    AskUser = False

    'Me!ctlCommonDialog.DialogTitle = "Locate Carnival File"
    'Me!ctlCommonDialog.Filter = "Carnival Files (*.mdb)|*.mdb|All (*.*)|*.*"
    'Me!ctlCommonDialog.DefaultExt = "mdb"
    'Me!ctlCommonDialog.FileName = ""
    
    'Me!ctlCommonDialog.ShowOpen
    
    Dim strFilter As String
    Dim lngFlags As Long
    strFilter = ahtAddFilterItem(strFilter, "Carnival Files (*.mdb, *.accdb)", "*.mdb; *.accdb")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    Result = ahtCommonFileOpenSave(InitialDir:="", Filter:=strFilter, FilterIndex:=0, Flags:=lngFlags, DialogTitle:="Locate Carnival File")
    
    'Result = Me!ctlCommonDialog.FileName
    ' Result = GetFileName("Select Database File", "Access Files (*.MDB)|*.MDB||", 1, "*")
    
    If (Result <> "") Then
        Result = Trim$(Result)
        If InStr(ReverseString(CStr(Result)), "\") <> 0 Then
            Set MyWS = DBEngine.Workspaces(0)
            Set MyDb = MyWS.Databases(0)
            Set Db = MyWS.OpenDatabase(Result)
            Set ITable = MyDb.OpenRecordset("SELECT * FROM [Inventory Attached Tables] Where [IF ID] = 2;")
            Do Until ITable.EOF
                Set TT = MyDb.TableDefs(ITable![Table Name])
                ITable.MoveNext
            Loop
            ITable.Close
            OldDB = DBPath()
            ''If Left$(Result, Len(OldDB)) = OldDB Then
            ''    NewDir = Right$(Result, Len(Result) - Len(OldDB))
            ''    If InStr(ReverseString(NewDir), "\") = 0 Then
            ''        NextCarn = NewDir
            ''        NewDir = ""
            ''    Else
            ''        NextCarn = Right$(NewDir, InStr(ReverseString(NewDir), "\") - 1)
            ''        NewDir = Left$(NewDir, Len(NewDir) - Len(NextCarn))
            ''    End If

            NextCarn = GetCarnivalFile(Result)
            NewDir = GetCarnivalRelDir(Result)
            
                DoCmd.SetWarnings False
                DoCmd.RunSQL "UPDATE DISTINCTROW Carnivals SET Carnivals.Filename = """ & NextCarn & """, Carnivals.[Relative Directory] = """ & NewDir & """, Carnivals.Available = True WHERE ((Carnivals.Carnival=""" & Me.List & """));"
                DoCmd.SetWarnings True
                AskUser = True
            ''Else
            ''    MsgBox "This carnival must be put in the global directory before it can be connected.", vbExclamation, "Message"
            ''End If
        End If
    End If
Exit_locateCarnival:
    DoCmd.SetWarnings True
    Exit Sub
Err_locateCarnival:
    MsgBox "This file does not meet the specified format for a carnival file.", vbExclamation, "Message"
    Resume Exit_locateCarnival

End Sub

Private Sub MakeActive_Click()
    
    Dim c As Variant

    List_DblClick (c)
    
End Sub

Private Sub New_Click()
    
    DoCmd.OpenForm "Carnival Copy", , , , , acDialog, "NEW"
    Me.List.Requery

End Sub

Private Sub Rename_Click()

    On Error Resume Next
    If IsNull(Me.List) Then
        MsgBox "Select a carnival to rename before pressing this button.", vbExclamation, "Message"
    Else
        GlobalVariable = Me!List.Column(0)
        DoCmd.OpenForm "Carnival Copy", , , , , acDialog, "RENAME"
        If IsNull(DLookup("[Carnival]", "Carnivals", "[Carnival] = """ & Me.List & """")) Then
            Form_Load
        End If
    End If

End Sub
Private Sub FinaliseCarnivalSelection()

    Dim Db As Database, rs As Recordset
    
    ' Check that there is at least one record in MiscHTML table
    If DCount("[TemplateFileSummary]", "MiscHTML") = 0 Then
        Set Db = CurrentDb
        Set rs = Db.OpenRecordset("MiscHTML")
        rs.AddNew
        rs!GenerateHTML = False
        rs.Update
        rs.Close
    End If
End Sub
Private Sub CompactCarnivalBut_Click()

    Dim fileName As Variant, Db As Database, FilePath As Variant, Response As Variant, TempName As Variant

    On Error GoTo Err_CompactCarnivalBut_Click

  If IsNull(Me!List) Then
    MsgBox ("You must select a carnival to compact.")
  ElseIf DLookup("[Available]", "Carnivals", "[Carnival] = """ & Me.List & """") Then

  
    PleaseWaitMsg = "Compacting carnival: " & Me!List & ".  Please wait ..."
    DoCmd.RunMacro "ShowPleaseWait"
    FilePath = CarnivalDir(DLookup("[Relative Directory]", "Carnivals", "[Carnival] = """ & Me.List & """"))
    fileName = FilePath & DLookup("[Filename]", "Carnivals", "[Carnival] = """ & Me.List & """")
    
    ' Check .mdb or .accdb
    If StrConv(Right(fileName, 6), vbLowerCase) = ".accdb" Then
      TempName = "__temp__.accdb"
    Else
      TempName = "__temp__.mdb"
    End If
    
    If FileExists(FilePath & TempName) Then Kill (FilePath & TempName)
    ReturnVar = SysCmd(acSysCmdSetStatus, "Compacting and Verifying database ...")
    DBEngine.CompactDatabase fileName, FilePath & TempName
    If FileExists(FilePath & TempName) Then
      If FileExists(fileName & ".old") Then
        Response = MsgBox("INFORMATIONAL ALERT ONLY: The compact action makes a new copy of the carnival file and appends .OLD to the original carnival file.  However an OLD carnival file already exists (usually because the compact action has been run on this carnival before).  Do you want to delete the OLD carnival and finish compacting the carnival?", vbYesNo + vbInformation + vbDefaultButton2, "Delete Old Carnival File")
        If Response = 6 Then
          Kill (fileName & ".old")
        Else
          Kill (FilePath & TempName)
          GoTo Exit_CompactCarnivalBut_Click
          
        End If
      End If
      
      Name fileName As (fileName & ".old")
      Name (FilePath & TempName) As fileName
    End If
  Else
    Response = MsgBox("The file for the selected carnival cannot be located.  Please locate the file manually and then retry compacting the carnival.", vbInformation)
    
  End If
  
  'acSysCmdAccessDir
  

Exit_CompactCarnivalBut_Click:
  DoCmd.RunMacro "ClosePleaseWait"
  ReturnVar = SysCmd(acSysCmdClearStatus)
    Exit Sub

Err_CompactCarnivalBut_Click:
    MsgBox Err.Description
    Resume Exit_CompactCarnivalBut_Click
    
End Sub
