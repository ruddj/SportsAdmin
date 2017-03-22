Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =5782
    ItemSuffix =50
    Left =1020
    Top =4695
    Right =4155
    Bottom =8655
    OrderBy ="EventTypeSub1.Age"
    RecSrcDt = Begin
        0xa54ca5b911cde140
    End
    RecordSource ="SELECT DISTINCTROW Events.E_Code, Events.ET_Code, Events.Age, Events.Sex, Events"
        ".Include, Events.nRecord, Events.Record, Events.RecName, Events.RecHouse FROM Ev"
        "ents ORDER BY Events.Age, Events.Sex;"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =255
    AllowLayoutView =0
    Begin
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =235
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =3317
                    Width =680
                    Height =227
                    Name ="E_Code"
                    ControlSource ="E_Code"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =455
                    Height =235
                    Name ="Text9"
                    Caption ="Age"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =638
                    Width =453
                    Height =220
                    Name ="Text8"
                    Caption ="Sex"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2112
                    Width =453
                    Height =220
                    Name ="Text33"
                    Caption ="Inc."
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1319
                    Width =600
                    Height =225
                    Name ="Text39"
                    Caption ="Record"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =291
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =45
                    Width =620
                    Height =245
                    ColumnWidth =750
                    ColumnOrder =0
                    Name ="Age"
                    ControlSource ="Age"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Age Division (eg. 13_U, 14, 15, 16, 17_U, OPEN)"

                End
                Begin CheckBox
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =2295
                    Top =45
                    Width =197
                    Height =215
                    TabIndex =3
                    Name ="Include"
                    ControlSource ="Include"
                    DefaultValue ="Yes"
                    ControlTipText ="Tick this if you want to include the event in the carnival."

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ListWidth =567
                    Left =705
                    Width =565
                    Height =245
                    ColumnWidth =585
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Sex"
                    ControlSource ="Sex"
                    RowSourceType ="Value List"
                    RowSource ="M;F;-"
                    ColumnWidths ="567"
                    FontName ="Arial"
                    ControlTipText ="Division gender."

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    BackStyle =0
                    Left =1344
                    Width =860
                    Height =253
                    TabIndex =2
                    Name ="Record"
                    ControlSource ="Record"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Double-Click to edit the records for this division."

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =3146
                    Width =785
                    Height =245
                    TabIndex =4
                    Name ="nRecord"
                    ControlSource ="nRecord"
                    FontName ="Arial"

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =4054
                    Width =785
                    Height =245
                    TabIndex =5
                    Name ="RecHouse"
                    ControlSource ="RecHouse"
                    FontName ="Arial"

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =4905
                    Width =785
                    Height =245
                    TabIndex =6
                    Name ="RecName"
                    ControlSource ="RecName"
                    FontName ="Arial"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2535
                    Top =15
                    Width =336
                    Height =276
                    TabIndex =7
                    Name ="DeleteDivisionBut"
                    Caption ="Command48"
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
                        0x00000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =623
            BackColor =-2147483633
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =45
                    Top =60
                    Width =2790
                    Height =285
                    FontSize =8
                    FontWeight =400
                    Name ="EditRecord"
                    Caption ="Edit Record For Selected Division"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Edit the record for the selected division."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =75
                    Top =375
                    Width =2730
                    Height =210
                    FontWeight =700
                    ForeColor =255
                    Name ="Alert"
                    Caption ="ALERT: No Divisions are setup"
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

Private Sub Button16_Click()
On Error GoTo Err_Button16_Click


    DoCmd.GoToRecord , , A_NEXT

Exit_Button16_Click:
    Exit Sub

Err_Button16_Click:
    MsgBox Error$
    Resume Exit_Button16_Click
    
End Sub

Private Sub Button17_Click()
On Error GoTo Err_Button17_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Button17_Click:
    Exit Sub

Err_Button17_Click:
    MsgBox Error$
    Resume Exit_Button17_Click
    
End Sub

Private Sub Button36_Click()

End Sub

Private Sub Button43_Click()

End Sub

Private Sub Button45_Click()

End Sub

Private Sub Button47_Click()
On Error GoTo Err_Button47_Click


    

Exit_Button47_Click:
    Exit Sub

Err_Button47_Click:
    MsgBox Error$
    Resume Exit_Button47_Click
    
End Sub

Private Sub EditRec_Click()

    'DoCmd.RunCommand acCmdSaveRecord
    
End Sub

Private Sub Age_AfterUpdate()

  DoUpdateEventCompetitorAge = True
  
End Sub

Private Sub EditRecord_Click()
On Error GoTo Err_EditRecord_Click


    Record_DblClick (Cancel)

Exit_EditRecord_Click:
    Exit Sub

Err_EditRecord_Click:
    MsgBox Error$
    Resume Exit_EditRecord_Click
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

    If IsNull(Me![Age]) Then
        Response = MsgBox("You must enter an age group.", vbInformation)
        Cancel = True
    ElseIf IsNull(Me![Sex]) Then
        Response = MsgBox("You must enter a valid gender.", vbInformation)
        Cancel = True
    End If

End Sub

Private Sub Form_Current()

  If NoFormRecords(Me.RecordsetClone) Then
    Me!Alert.visible = True
  Else
    Me!Alert.visible = False
  End If
  
End Sub

Private Sub Record_DblClick(Cancel As Integer)
    
    On Error GoTo Err_Record_DblClick
    
    DoCmd.RunCommand acCmdSaveRecord

    If Not IsNull(Me![E_Code]) And Not IsNull(Me![Age]) Then
        
        DoCmd.RunCommand acCmdSaveRecord

        DoCmd.OpenForm "EventRecord", , , "[E_Code]=" & Me![E_Code], , acDialog
        'Me![Record] = 45
        'Me![nRecord] = 45
    
        RecDate = DLookup("[MaxOfDate]", "Record-Most Recent", "[E_Code]=" & Me![E_Code])
        If Not IsNull(RecDate) Then
            DoCmd.RunCommand acCmdSaveRecord
            Criteria = "[E_Code]=" & Me![E_Code] & " AND [Date]=#" & Format(RecDate, "mm/dd/yy") & "#"
            res = DLookup("[Result]", "Records", Criteria)
            nRes = DLookup("[nResult]", "Records", Criteria)
            Gname = DLookup("[Gname]", "Records", Criteria)
            Sname = DLookup("[Surname]", "Records", Criteria)
            Hcode = DLookup("[H_code]", "Records", Criteria)
            Hid = DLookup("[H_id]", "House", "[H_Code]=""" & Hcode & """")
            
            Me![Record] = res
            Me![nRecord] = nRes
            Me![RecName] = Trim(UCase(Sname)) & ", " & Trim(Gname)
            Me![RecHouse] = Hid
            
            'q = "UPDATE DISTINCTROW Events SET "
            'q = q & " Events.nRecord=" & nRes & ", Events.Record=""" & Res & """, Events.RecName=""" & RecName & """, Events.RecHouse=" & Hid
            'q = q & " where Events.E_Code=" & Me![E_Code]
            
            'DoCmd SetWarnings False
            'DoCmd RunSQL q
            'DoCmd SetWarnings True
        Else
            Me![Record] = Null
            Me![nRecord] = Null
            Me![RecName] = Null
            Me![RecHouse] = Null
        End If
    End If

Exit_Record_DblClick:
    Exit Sub
    
Err_Record_DblClick:
    MsgBox Error
    GoTo Exit_Record_DblClick
End Sub

Private Sub DeleteDivisionBut_Click()
On Error GoTo Err_DeleteDivisionBut_Click

  Response = MsgBox("Are you sure you want to delete this division?  All records for this division will be deleted also.", vbYesNo + vbInformation + vbDefaultButton2, "Confirm division delete.")
  If Response = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
  End If

Exit_DeleteDivisionBut_Click:
    Exit Sub

Err_DeleteDivisionBut_Click:
    MsgBox Err.Description
    Resume Exit_DeleteDivisionBut_Click
    
End Sub
