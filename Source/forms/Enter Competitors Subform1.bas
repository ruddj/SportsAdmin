Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =8991
    ItemSuffix =71
    Left =405
    Top =3015
    Right =14100
    Bottom =9705
    HelpContextId =110
    AfterDelConfirm ="[Event Procedure]"
    OrderBy ="EnterCompetitorsSF.Place"
    RecSrcDt = Begin
        0x914020067125e240
    End
    RecordSource ="EnterCompetitorsSF"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000292200000001000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    FilterOnLoad =255
    AllowLayoutView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
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
        Begin OptionButton
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =187
            Height =187
            LabelX =-236
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =187
            Height =187
            LabelX =-236
        End
        Begin OptionGroup
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            OldBorderStyle =0
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
            BackColor =12632256
        End
        Begin ListBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
            BackColor =12632256
        End
        Begin ComboBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
            BackColor =12632256
        End
        Begin FormHeader
            Height =340
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4305
                    Top =72
                    Width =543
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text20"
                    Caption ="Place"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Top =72
                    Width =495
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="LaneTXT"
                    Caption ="Lane"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1944
                    Top =72
                    Width =2147
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text24"
                    Caption ="Name"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4968
                    Top =72
                    Width =843
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text30"
                    Caption ="Result"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =504
                    Top =72
                    Width =705
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text34"
                    Caption ="Team"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =7426
                    Top =72
                    Width =454
                    Height =213
                    ColumnOrder =0
                    BackColor =-2147483633
                    Name ="nRes"
                    ControlSource ="nResult"
                    FontName ="Tahoma"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1368
                    Top =72
                    Width =360
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text66"
                    Caption ="Pts"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =8192
                    Width =621
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Count"
                    ControlSource ="=Count([H_Code])"
                    FontName ="Tahoma"

                    LayoutCachedLeft =8192
                    LayoutCachedWidth =8813
                    LayoutCachedHeight =240
                End
            End
        End
        Begin Section
            Height =311
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =564
                    Width =795
                    Height =285
                    TabIndex =1
                    BackColor =-2147483633
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    Format =">"
                    StatusBarText ="Personal ID Number"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =6041
                    Width =527
                    Height =285
                    FontSize =7
                    TabIndex =6
                    BackColor =-2147483633
                    Name ="Unit"
                    ControlSource ="=[Forms]![EnterCompetitors]![Units]"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4290
                    Width =645
                    Height =285
                    TabIndex =4
                    BackColor =16777215
                    Name ="Place"
                    ControlSource ="Place"
                    StatusBarText ="Place gained by competitor"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Enter the place gained by the competitor."
                    HorizontalAnchor =1

                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =405
                    Height =285
                    BackColor =16777215
                    Name ="Lane"
                    ControlSource ="Lane"
                    StatusBarText ="Lane: Lane 0 defaults to the lane allocated to the competitors House"
                    FontName ="Tahoma"
                    ControlTipText ="The lane will be added automatically if you have allocated lanes to teams.  Othe"
                        "rwise enter the lane manually."

                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =5
                    ListWidth =5103
                    Left =1944
                    Width =2320
                    Height =286
                    TabIndex =3
                    BoundColumn =3
                    BackColor =16777215
                    Name ="Fname"
                    ControlSource ="PIN"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="2612;680;398;680;455"
                    StatusBarText ="Choose a competitor from the list"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnNotInList ="[Event Procedure]"
                    ControlTipText ="Select the competitors that are to compete in this event."
                    HorizontalAnchor =2

                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    Left =4968
                    Width =1047
                    Height =286
                    TabIndex =5
                    BackColor =16777215
                    Name ="Res"
                    ControlSource ="Result"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Enter the result gained by the competitor."
                    HorizontalAnchor =1

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =8844
                    Width =147
                    Height =286
                    TabIndex =13
                    BackColor =16777215
                    Name ="nResult"
                    ControlSource ="nResult"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8844
                    LayoutCachedWidth =8991
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =7676
                    Width =327
                    Height =286
                    TabIndex =7
                    BackColor =16777215
                    Name ="F_Lev"
                    ControlSource ="F_Lev"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8073
                    Width =177
                    Height =286
                    TabIndex =8
                    BackColor =16777215
                    Name ="HE_Code"
                    ControlSource ="HE_Code"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8073
                    LayoutCachedWidth =8250
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8277
                    Width =132
                    Height =286
                    TabIndex =9
                    BackColor =16777215
                    Name ="E_Code"
                    ControlSource ="E_Code"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8277
                    LayoutCachedWidth =8409
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =8447
                    Width =117
                    Height =286
                    TabIndex =10
                    BackColor =16777215
                    Name ="Heat"
                    ControlSource ="Heat"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8447
                    LayoutCachedWidth =8564
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =8588
                    Width =117
                    Height =286
                    TabIndex =11
                    BackColor =16777215
                    Name ="H_ID"
                    ControlSource ="H_ID"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8588
                    LayoutCachedWidth =8705
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =7256
                    Width =765
                    Height =285
                    TabIndex =14
                    Name ="PIN"
                    ControlSource ="PIN"
                    StatusBarText ="Personal ID Number"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =7256
                    LayoutCachedWidth =8021
                    LayoutCachedHeight =285
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6600
                    Top =15
                    Width =314
                    Height =286
                    FontSize =8
                    FontWeight =400
                    TabIndex =15
                    Name ="Memo"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Relay Team Members / Competitor comment."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =1
                    Left =8702
                    Width =117
                    Height =286
                    TabIndex =12
                    BackColor =16777215
                    Name ="MemoFld"
                    ControlSource ="Memo"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8702
                    LayoutCachedWidth =8819
                    LayoutCachedHeight =286
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1445
                    Width =420
                    Height =285
                    TabIndex =2
                    BackColor =-2147483633
                    Name ="Points"
                    ControlSource ="Points"
                    StatusBarText ="Points gained by the competitor.  This can be edited manuallu but is normally up"
                        "dated automatically depending upon the place gained."
                    FontName ="Tahoma"
                    ControlTipText ="Points gained by the competitor.  This can be edited manuallu but is normally up"
                        "dated automatically depending upon the place gained."

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6930
                    Top =15
                    Width =284
                    Height =286
                    TabIndex =16
                    Name ="DeleteCompetitorBut"
                    Caption ="Command70"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadadadadadaadad00adad00adaddadad00ad00adada ,
                        0xadadad0000adadaddadadad00adadadaadadad0000adadaddadad00ad00adada ,
                        0xadad00adad00adaddadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Competitor"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter2"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons

Dim ShowAllCompetitors As Variant

Private Sub Fname_BeforeUpdate(Cancel As Integer)

    x = 1
    
End Sub

Private Sub Fname_DblClick(Cancel As Integer)
    
    If Not IsNull(Me![PIN]) Then
        Call MaintainCompetitor("EDIT", Me!PIN)
        Me![PIN].Requery
    End If

End Sub

Private Sub Fname_NotInList(NewData As String, Response As Integer)
' Prompt to add competitor to database
' Check that there is no competitor with similar name
  
  Dim s As String
  
  Response = MsgBox("The name you have entered does not exist in the list.  Do you want to add this person to the carnival?", vbQuestion + vbYesNo + vbDefaultButton2)
  If Response = vbNo Then Exit Sub
  
  s = Trim(NewData)
  s = s & "|" & Trim(Me.Parent.Age)
  s = s & "|" & Trim(Me.Parent.Sex)
  
  GlobalCancel = True
  DoCmd.OpenForm "Competitor-QuickAdd", , , , , acDialog, s
  
  If GlobalCancel Then
    Response = acDataErrContinue
  Else
    Me.Fname.Undo
    Me.Fname.Requery
    Me.Fname.Text = GlobalVariable
    'Response = acDataErrAdded
    'Me.Fname.Text = GlobalVariable
  End If


End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)

    Call Update_Lane_Assignments(Forms![EnterCompetitors]![E_Code], Forms![EnterCompetitors]![F_Lev], Forms![EnterCompetitors]![Heat])
    

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Form_BeforeUpdate_Err

  If IsNull(Me!PIN) Or Me!PIN = 0 Then
    Response = MsgBox("You must select a competitor from the list (or push Esc to cancel).", vbInformation)
    Cancel = True
  Else
    'If IsNull(Me![Place]) Then
    '    Me![Place] = 0
    'End If
    If Me![Place] = 0 Then
        Me![Place] = Null
    End If
        
    If Forms![EnterCompetitors]![Lane_Cnt] = 0 Then
        Me![Lane] = 0
    ElseIf (Me![Lane] = 0) Then
        Me![Lane] = Calculate_Competitor_Lane(Me![E_Code], Me![F_Lev], Me![H_ID], Me![Heat])
    End If
  End If
    
Form_BeforeUpdate_Exit:
  Exit Sub
  
Form_BeforeUpdate_Err:
  MsgBox "An error has occurred in [Form_BeforeUpdate]: " & Err.Description, vbCritical
  
End Sub

Private Sub Form_Open(Cancel As Integer)

    ShowAllCompetitors = False

End Sub

Private Sub Memo_Click()
On Error GoTo Memo_Click_Err
    
    Q = "UPDATE DISTINCTROW [Temporary Memo] SET [Temporary Memo].[Memo] = """ & Me![MemoFld] & """"
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True

    GlobalCancel = False
    DoCmd.OpenForm "EnterCompetitor Memo", , , , , acDialog, Me![Fname].RowSource
    
    MemoValue = DLookup("[Memo]", "Temporary Memo")
    If Not GlobalCancel Then
        Me![MemoFld] = MemoValue
    End If
    
Memo_Click_Exit:
  Exit Sub

Memo_Click_Err:
  MsgBox "An error has occurred in [Memo_Click]: " & Err.Description, vbCritical
  
End Sub

Private Sub Place_AfterUpdate()

  GlobalPlaceChange = True
  GlobalChange = True
  If IsNull(Me![Place]) Then
      Me![Points] = Null
  Else
      Me![Points] = DLookup("[Points]", "PointsScale", "[Place]=" & Me![Place] & " AND [PtScale]=""" & Forms![EnterCompetitors]![PtScale] & """")
      If IsNull(Me![Points]) Then
          Me![Points] = 0
      End If
  End If

End Sub


Private Sub DeleteCompetitorBut_Click()
On Error GoTo Err_DeleteCompetitorBut_Click

  Response = MsgBox("Are you sure you want to delete this competitor?", vbYesNo + vbInformation + vbDefaultButton2, "Confirm delete competitor")
  If Response = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
  End If
  
Exit_DeleteCompetitorBut_Click:
    Exit Sub

Err_DeleteCompetitorBut_Click:
    MsgBox Err.Description
    Resume Exit_DeleteCompetitorBut_Click
    
End Sub


Private Sub Res_AfterUpdate()
On ERRROR GoTo Res_AfterUpdate_Err

  GlobalChange = True

    Dim res As String
    Dim Runit As String
    Dim cUnit As String
    Dim Valu As String
    Dim nRes As Double
    Dim Power As Integer
    Dim Delm As String, nValu As String
    Dim i As Integer
    Dim AddZero As Integer
    Dim success As Boolean
    
  If Not (IsNull(Me![res])) Then
    
    res = Me![res]
    Runit = [Forms]![EnterCompetitors]![Units]
    Call Calculate_Results(res, nValu, Runit, success)
    
    If success Then
      Forms![EnterCompetitors]![EC_Subform].Form![nRes] = res
      Forms![EnterCompetitors]![EC_Subform].Form![res] = nValu
    Else
      Me.res.SetFocus
    End If
  Else
    ' When the Result (time or distance or points) is set to NULL then
    ' set Numeric Result and Place to 0

    Me.[nResult] = 0
    Me.[Place] = 0

  End If

Res_AfterUpdate_Exit:
  Exit Sub
  
Res_AfterUpdate_Err:
  MsgBox "An error occurred in [Res_AfterUpdate]:" & Err.Description


End Sub

Private Sub Res_BeforeUpdate(Cancel As Integer)
On Error GoTo Res_BeforeUpdate_Err

    Dim res As String
    Dim Runit As String
    Dim nValu As String
    Dim success As Boolean
    
    res = Nz(Me![res])
    Runit = [Forms]![EnterCompetitors]![Units]
    Call Calculate_Results(res, nValu, Runit, success)
    
    Cancel = Not success
    
Res_BeforeUpdate_Exit:
  Exit Sub
  
Res_BeforeUpdate_Err:
  MsgBox "Are error has occurred in 'Res_BeforeUpdate': " & Err.Description, vbCritical
  GoTo Res_BeforeUpdate_Exit
End Sub
