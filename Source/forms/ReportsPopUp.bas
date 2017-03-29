Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3344
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =8505
    Top =2325
    Right =11760
    Bottom =6930
    TimerInterval =1000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xef8ca22f10fae140
    End
    RecordSource ="SELECT [Misc-ReportsPopUp].* FROM [Misc-ReportsPopUp];"
    Caption ="Open Reports"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnTimer ="[Event Procedure]"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Tab
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3855
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ListBox
                    Visible = NotDefault
                    OverlapFlags =215
                    ColumnCount =2
                    Left =150
                    Top =135
                    Width =3004
                    Height =3635
                    Name ="ReportList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportList.ReportName, ReportList.ReportCaption FROM ReportList WHERE ((("
                        "ReportList.Open)=True)) ORDER BY ReportList.ReportCaption;"
                    ColumnWidths ="0;2948"
                    StatusBarText ="Select the report you want to look at."
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Select the report you want to look at."

                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =396
                    Top =907
                    Width =2438
                    Height =1134
                    FontWeight =700
                    Name ="Label4"
                    Caption ="Finding open reports ..."
                End
            End
        End
        Begin FormFooter
            Height =1474
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Left =60
                    Top =15
                    Width =3184
                    Height =1410
                    Name ="TabCtl17"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =195
                            Top =420
                            Width =2911
                            Height =870
                            Name ="Page18"
                            Caption ="Report Actions"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =202
                                    Top =422
                                    Width =1418
                                    Height =397
                                    FontWeight =700
                                    Name ="CloseRepBut"
                                    Caption ="Close Report"
                                    OnClick ="[Event Procedure]"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =195
                                    Top =887
                                    Width =1418
                                    Height =397
                                    Name ="CloseAllReportsBut"
                                    Caption ="Close All Reports"
                                    OnClick ="[Event Procedure]"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1688
                                    Top =420
                                    Width =1418
                                    Height =397
                                    FontWeight =700
                                    Name ="PrintReport"
                                    Caption ="Print Report"
                                    OnClick ="[Event Procedure]"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1688
                                    Top =885
                                    Width =1418
                                    Height =397
                                    Name ="PrintAllReportsBut"
                                    Caption ="Print  All Reports"
                                    OnClick ="[Event Procedure]"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =420
                            Width =2910
                            Height =870
                            Name ="Page19"
                            Caption ="Popup Position"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    Left =1296
                                    Top =675
                                    Width =891
                                    Name ="ReportPopupX"
                                    ControlSource ="ReportPopupX"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =3
                                            Left =331
                                            Top =675
                                            Width =915
                                            Height =240
                                            Name ="Label60"
                                            Caption ="X Position:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =223
                                    Left =1295
                                    Top =951
                                    Width =891
                                    TabIndex =1
                                    Name ="ReportPopupY"
                                    ControlSource ="ReportPopupY"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =3
                                            Left =330
                                            Top =951
                                            Width =915
                                            Height =240
                                            Name ="Label62"
                                            Caption ="Y Position:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =210
                                    Top =541
                                    Width =2836
                                    Height =731
                                    Name ="Box65"
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =515
                                    Top =457
                                    TabIndex =2
                                    Name ="Check63"

                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            Left =751
                                            Top =420
                                            Width =1500
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Label64"
                                            Caption ="Show Report Popup"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =2252
                                    Top =710
                                    Width =227
                                    Height =227
                                    TabIndex =3
                                    Name ="xm"
                                    Caption ="Command23"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000008000000080000000100180000000000c00000000000000000000000 ,
                                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff000000000000000000000000
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Use the space bar to push this button repeatedly."
                                    Picture ="minus.bmp"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =2479
                                    Top =710
                                    Width =227
                                    Height =227
                                    TabIndex =4
                                    Name ="xp"
                                    Caption ="Command23"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000008000000080000000100180000000000c00000000000000000000000 ,
                                        0x0000000000000000ffffffffffffffffff000000000000ffffffffffffffffff ,
                                        0xffffffffffffffffff000000000000ffffffffffffffffffffffffffffffffff ,
                                        0xff000000000000ffffffffffffffffff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0xffffffffffffffffff000000000000ffffffffffffffffffffffffffffffffff ,
                                        0xff000000000000ffffffffffffffffffffffffffffffffffff000000000000ff ,
                                        0xffffffffffffffff000000000000000000000000
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Use the space bar to push this button repeatedly."
                                    Picture ="plus.bmp"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =2250
                                    Top =945
                                    Width =227
                                    Height =227
                                    TabIndex =5
                                    Name ="ym"
                                    Caption ="Command23"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000008000000080000000100180000000000c00000000000000000000000 ,
                                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Use the space bar to push this button repeatedly."
                                    Picture ="minus.bmp"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =2477
                                    Top =945
                                    Width =227
                                    Height =227
                                    TabIndex =6
                                    Name ="yp"
                                    Caption ="Command23"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000008000000080000000100180000000000c00000000000000000000000 ,
                                        0x0000000000000000ffffffffffffffffff000000000000ffffffffffffffffff ,
                                        0xffffffffffffffffff000000000000ffffffffffffffffffffffffffffffffff ,
                                        0xff000000000000ffffffffffffffffff00000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0xffffffffffffffffff000000000000ffffffffffffffffffffffffffffffffff ,
                                        0xff000000000000ffffffffffffffffffffffffffffffffffff000000000000ff ,
                                        0xffffffffffffffff000000000000000000000000
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Use the space bar to push this button repeatedly."
                                    Picture ="plus.bmp"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                    End
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
Option Compare Database
Option Explicit

Dim TimerStarted  As Boolean

Private Sub CloseAllReportsBut_Click()
On Error Resume Next

  'DoCmd.Echo False
  Dim X As Integer
  
  For X = 0 To Reports.Count - 1
    DoCmd.SelectObject A_REPORT, Reports(0).Name, False
    DoCmd.Close
    
  Next X

  'DoCmd.Echo True

End Sub

Private Sub Form_Close()

'  DoCmd.RunCommand acCmdSaveLayout
  
End Sub

Private Sub Form_Load()

  TimerStarted = True
  Me.InsideHeight = Me.FormInsideHeight
  
  Call Form_Resize
  Call PositionForm
  

End Sub

Private Sub CloseRepBut_Click()
On Error GoTo Err_CloseRepBut_Click

  If Not (VarEmpty(Me!ReportList)) Then
    DoCmd.Close acReport, Me!ReportList
    Call CreateReportList
  End If

Exit_CloseRepBut_Click:
    Exit Sub

Err_CloseRepBut_Click:
    MsgBox Err.Description
    Resume Exit_CloseRepBut_Click
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

  If Me.InsideHeight < 2000 Then Me.InsideHeight = 2000
'  MsgBox (Me.ReportList.Height)
  Me!ReportList.Height = Me.InsideHeight - 1700
'  MsgBox (Me.ReportList.Height)
  Me.InsideWidth = 3250
  Me.FormInsideHeight = Me.InsideHeight

End Sub

Private Sub Form_Timer()
On Error Resume Next

  If TimerStarted Then
    Call CreateReportList(True)
    TimerStarted = False
    Me!ReportList.visible = True
    Me!ReportList = Me!ReportList.ItemData(0)
    DoCmd.SelectObject A_REPORT, Me!ReportList, False
    DoCmd.RunCommand acCmdPreviewTwoPages
'    Call View_AfterUpdate
  End If

End Sub

Private Sub PrintAllReportsBut_Click()

On Error GoTo PrintAllReportsBut_Click_Err
  Dim X As Integer, msg As String
  
  For X = 0 To Reports.Count - 1
    DoCmd.SelectObject A_REPORT, Reports(X).Name, False
    DoCmd.RunCommand acCmdPrint
  Next X

PrintAllReportsBut_Click_Exit:
  Exit Sub
  
PrintAllReportsBut_Click_Err:
  msg = ""
  If (Err.Number = 2501) Then
    If (X < (Reports.Count - 1)) Then
      msg = "You have cancelled printing report '"
    End If
  Else
    msg = "An error has occured printing '"
  End If
  If msg <> "" Then
    Response = MsgBox(msg & Reports(X).Caption & "'.  Do you wish to continue printing?", vbInformation + vbYesNo)
    If Response = vbNo Then
      Exit Sub
    Else
      Resume Next
    End If
    
  End If
  
  GoTo PrintAllReportsBut_Click_Exit
  
End Sub

Private Sub PrintReport_Click()

On Error GoTo PrintReport_Click_Err:

  If Not (VarEmpty(Me!ReportList)) Then
    DoCmd.SelectObject A_REPORT, Me!ReportList, False
    Call DisplayPrintDialog
  End If
  
PrintReport_Click_Err:
  Exit Sub
  
End Sub

Private Sub ReportList_Click()
  
On Error GoTo ReportList_Click_Err

  If Not (VarEmpty(Me!ReportList)) Then
'    If Me!ApplyToAll Then
'      Call View_AfterUpdate
'    Else
      DoCmd.SelectObject A_REPORT, Me!ReportList, False
'    End If
    DoCmd.SelectObject A_FORM, "ReportsPopup", False
    
  End If

ReportList_Click_Exit:
  Exit Sub
  
ReportList_Click_Err:
  Call CreateReportList

End Sub
Private Sub MaximiseReport_Click()
End Sub

Private Sub View_AfterUpdate_Notused()
On Error GoTo Err_View_AfterUpdate

  If Not (VarEmpty(Me!ReportList)) Then
    DoCmd.SelectObject A_REPORT, Me!ReportList, False
    If Me!View = 1001 Then
      DoCmd.Minimize
    ElseIf Me!View = 1002 Then
      DoCmd.Restore
    ElseIf Me!View = 1003 Then
      DoCmd.Maximize
    Else
      DoCmd.RunCommand Me!View
    End If
  End If
  
Exit_View_AfterUpdate:
    Exit Sub

Err_View_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_View_AfterUpdate
    


End Sub
Private Sub Position_Click()
On Error GoTo Err_Position_Click


Exit_Position_Click:
    Exit Sub

Err_Position_Click:
    MsgBox Err.Description
    Resume Exit_Position_Click
    
End Sub

Private Sub ReportPopupX_AfterUpdate()

  Call PositionForm
  
End Sub

Private Sub ReportPopupY_AfterUpdate()

  Call PositionForm

End Sub

Private Sub PositionForm()

On Error GoTo PositionForm_Exit
    
  Dim X As Integer, Y As Integer
  
  If VarEmpty(Me!ReportPopupX) Then
    Me!ReportPopupX = 0
  Else
    If Me!ReportPopupX < 0 Then
      Me!ReportPopupX = 0
    ElseIf Me!ReportPopupX > 11500 Then
      Me!ReportPopupX = 11500
    End If
  End If
  
  X = Me!ReportPopupX
    
  If VarEmpty(Me!ReportPopupY) Then
    Me!ReportPopupY = 0
  Else
    If Me!ReportPopupY < 0 Then
      Me!ReportPopupY = 0
    ElseIf Me!ReportPopupY > 9000 Then
      Me!ReportPopupY = 9000
    End If
  End If
  Y = Me!ReportPopupY
  
  DoCmd.MoveSize X, Y
  
PositionForm_Exit:

End Sub

Private Sub Command20_Click()
On Error GoTo Err_Command20_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command20_Click:
    Exit Sub

Err_Command20_Click:
    MsgBox Err.Description
    Resume Exit_Command20_Click
    
End Sub
Private Sub Command23_Click()
On Error GoTo Err_Command23_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command23_Click:
    Exit Sub

Err_Command23_Click:
    MsgBox Err.Description
    Resume Exit_Command23_Click
    
End Sub

Private Sub xm_Click()

  Me!ReportPopupX = Me!ReportPopupX - 500
  Call PositionForm
  
End Sub

Private Sub xp_Click()

  Me!ReportPopupX = Me!ReportPopupX + 500
  Call PositionForm

End Sub

Private Sub ym_Click()

  Me!ReportPopupY = Me!ReportPopupY - 500
  Call PositionForm


End Sub

Private Sub yp_Click()

  Me!ReportPopupY = Me!ReportPopupY + 500
  Call PositionForm


End Sub
