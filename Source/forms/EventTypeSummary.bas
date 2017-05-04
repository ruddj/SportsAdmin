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
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =10374
    ItemSuffix =9
    Left =-20685
    Top =7095
    Right =-8490
    Bottom =14940
    HelpContextId =60
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xb2463d042dc7e140
    End
    Caption ="Event Summary"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnResize ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin Section
            Height =6746
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    ColumnCount =6
                    Left =265
                    Top =343
                    Width =8440
                    Height =6260
                    BorderColor =12632256
                    Name ="Summary"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT EventType.ET_Code, EventType.ET_Des AS Event, EventType.Units, EventType."
                        "Lane_Cnt AS Lanes, ReportTypes.Desc AS [Report Type], EventType.Include FROM Rep"
                        "ortTypes INNER JOIN EventType ON ReportTypes.R_Code = EventType.R_Code ORDER BY "
                        "EventType.ET_Des;"
                    ColumnWidths ="0;2552;680;680;3686;466"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    ControlTipText ="Double-Click an event to edit it."
                    VerticalAnchor =2

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =255
                            Top =60
                            Width =1920
                            Height =240
                            FontWeight =700
                            Name ="Text1"
                            Caption ="Summary of Events"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8992
                    Top =6095
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =1
                    Name ="CloseBut"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8985
                    Top =1553
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    ForeColor =128
                    Name ="DeleteBut"
                    Caption ="Delete Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8985
                    Top =2175
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =32768
                    Name ="CopyBut"
                    Caption ="Copy Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Copy the select event."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8985
                    Top =330
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    ForeColor =32768
                    Name ="AddBut"
                    Caption ="Add Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Add a new event."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8990
                    Top =5410
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    HelpContextId =60
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8985
                    Top =945
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    ForeColor =8404992
                    Name ="Edit"
                    Caption ="Edit Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Edit the selected event."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8990
                    Top =3400
                    Width =1134
                    Height =765
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    ForeColor =8404992
                    Name ="UpdateRecords"
                    Caption ="Update Event Records"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Maintenance: Check if the record for selected event has been broken."

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

' Form Dimensions
Dim lMinHeight As Long
Dim lMinWidth As Long

Private Sub AddBut_Click()
On Error GoTo Err_AddBut_Click
    
    GlobalCancel = True
    DoCmd.OpenForm "EventTypeCopy", , , , , acDialog, "ADD"
    'Stop
    If Not GlobalCancel Then
        
        [Summary].Requery
        Me![Summary] = GlobalVariable
        If Not IsEmpty(Me![Summary]) Then
          DoCmd.OpenForm "EventType-Wizard", , , "[ET_Code] = " & [Summary], , acDialog, "EDIT"
          Call EditEvent
        End If

    End If

Exit_AddBut_Click:
    Exit Sub

Err_AddBut_Click:
    MsgBox Error$
    Resume Exit_AddBut_Click
    
End Sub

Private Sub CloseBut_Click()
On Error GoTo Err_CloseBut_Click


    DoCmd.Close

Exit_CloseBut_Click:
    Exit Sub

Err_CloseBut_Click:
    MsgBox Error$
    Resume Exit_CloseBut_Click
    
End Sub

Private Sub CopyBut_Click()
On Error GoTo Err_CopyBut_Click

    If IsNull([Summary]) Then
        MsgBox ("You must select an event to copy.")
    Else
        DoCmd.OpenForm "EventTypeCopy", , , , , acDialog, "COPY"

        [Summary].Requery

    End If

Exit_CopyBut_Click:
    Exit Sub

Err_CopyBut_Click:
    MsgBox Error$
    Resume Exit_CopyBut_Click
    
End Sub

Private Sub DeleteBut_Click()
On Error GoTo Err_DeleteBut_Click

    ' Generate Warning - # Competitors, Records,
 If IsNull([Summary]) Then
    MsgBox ("You must select an event to delete.")
 Else
    
    NumCompetitors = DCount("[ET_Code]", "EventTypeCompetitors", "[Et_Code] = " & [Summary])
    NumRecords = DCount("[ET_Code]", "EventTypeRecords", "[Et_Code] = " & [Summary])
    NumHeats = DCount("[ET_Code]", "EventTypeHeats", "[Et_Code] = " & [Summary])

    WarningMessage = "There are presently " & NumCompetitors & " competitors, " & NumRecords & " record(s) and " & NumHeats & " heat(s) connected to this event type.  If you continue with this delete operation, all this data will be lost.  Do you wish to continue?"

    Response = MsgBox(WarningMessage, vbYesNo + vbCritical)
        
    If Response = vbYes Then 'Yes
        Q = "DELETE DISTINCTROW EventType.ET_Code FROM EventType WHERE EventType.ET_Code= " & [Summary]
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True
        [Summary].Requery
        
    End If

  End If

Exit_DeleteBut_Click:
    Exit Sub

Err_DeleteBut_Click:
    MsgBox Error$
    Resume Exit_DeleteBut_Click
    
End Sub

Private Sub Edit_Click()
    
    Summary_DblClick (Cancel)

End Sub

Private Sub Form_Open(Cancel As Integer)
    lMinHeight = frmHeight(Me)
    lMinWidth = Me.Width
    
    DoUpdateEventCompetitorAge = False
  
End Sub

Private Sub Form_Resize()
    If Not m_blResize Then Call glrMinWindowSize(Me, lMinHeight, lMinWidth, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Err

  If DoUpdateEventCompetitorAge Then Call UpdateEventCompetitorAge
  
Form_Unload_Exit:
  On Error Resume Next
  Exit Sub

Form_Unload_Err:
  Call DisplayErrMsg("Form_Unload")
  Resume Form_Unload_Exit

End Sub

Private Sub Summary_DblClick(Cancel As Integer)

On Error GoTo err_sdc

  Call EditEvent
    
exit_sdc:
    Exit Sub

err_sdc:
    MsgBox (Error$)
    GoTo exit_sdc

End Sub

Private Sub EditEvent()

    If IsNull(Me![Summary]) Then
        MsgBox ("Select an event to edit.")
    Else
        DoCmd.OpenForm "EventType", , , "[ET_Code] = " & [Summary], , acDialog, "EDIT"
    End If

End Sub
Private Sub Summary_KeyDown(KeyCode As Integer, Shift As Integer)

    'Stop
    If KeyCode = 46 Then
        DeleteBut_Click
    ElseIf KeyCode = 13 Then
        Summary_DblClick (Cancel)
    End If
        
End Sub

Private Sub UpdateRecords_Click()

On Error GoTo UpdateRecords_Click_Err

  Dim Response As Variant
  
  Response = MsgBox("This action will update the records for all events.  Do you want to continue?", vbQuestion + vbYesNo + vbDefaultButton2)
  If Response = vbYes Then
    If DCount("[Heat]", "Heats") > 0 Then
        PleaseWaitMsg = "Updating event records ..."
        DoCmd.RunMacro "ShowPleaseWait"
        Dim db As Database, Rs As Recordset
        'Stop
        Set db = DBEngine.Workspaces(0).Databases(0)
        Set Rs = db.OpenRecordset("Heats", dbOpenDynaset)   ' Create dynaset.
        GlobalVariable = False
        Rs.MoveLast
        Tot = Rs.RecordCount
        msg = "Updating Event Records ..."
        ReturnValue = SysCmd(acSysCmdInitMeter, msg, Tot)    ' Display message in status bar.
        X = 1
                
        Rs.MoveFirst
        While Not Rs.EOF
            Call CheckIfRecordBroken(Rs!E_Code, -1, -1)
            Rs.MoveNext
            ReturnValue = SysCmd(acSysCmdUpdateMeter, X)   ' Update meter.
            X = X + 1
        Wend
    
        Rs.Close
        ReturnValue = SysCmd(acSysCmdRemoveMeter)   ' Update meter.
    End If
  End If
  
UpdateRecords_Click_Exit:
  DoCmd.RunMacro "ClosePleaseWait"
  Exit Sub
  
UpdateRecords_Click_Err:
  MsgBox ("An error has occured in [UpdateRecords_Click]: " & Err.Description)
  GoTo UpdateRecords_Click_Exit
  
End Sub
