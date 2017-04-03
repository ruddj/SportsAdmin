Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    GridX =20
    GridY =20
    Width =10368
    ItemSuffix =30
    Right =14340
    Bottom =9495
    HelpContextId =40
    RecSrcDt = Begin
        0x637d3e042dc7e140
    End
    RecordSource ="Miscellaneous"
    Caption ="Team Summary"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Width =1701
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
        Begin Tab
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =6973
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =9000
                    Top =6310
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    Name ="Close"
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
                    AccessKey =84
                    TextFontFamily =34
                    Left =9000
                    Top =1112
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    ForeColor =32768
                    Name ="AddB"
                    Caption ="Add &Team"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =76
                    TextFontFamily =34
                    Left =9000
                    Top =1792
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    ForeColor =128
                    Name ="DeleteB"
                    Caption ="De&lete  Team"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Delete the selected team."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =1690
                    Left =8928
                    Top =4112
                    Width =805
                    Height =227
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="HT_Code"
                    ControlSource ="HouseType"
                    RowSourceType ="Table/Query"
                    RowSource ="Select [HT_Code],[Desc] From [HouseTypes];"
                    ColumnWidths ="0;1440"
                    FontName ="Tahoma"
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    Left =8928
                    Top =3886
                    Width =870
                    Height =225
                    Name ="Text15"
                    Caption ="Show the following:"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =72
                    TextFontFamily =34
                    Left =9000
                    Top =4935
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    HelpContextId =40
                    Name ="Help"
                    Caption ="&Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    TextFontFamily =34
                    Left =9000
                    Top =2520
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =8404992
                    Name ="Extra"
                    Caption ="&Allocate Extra Points"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    TextFontFamily =34
                    Left =9000
                    Top =3153
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    ForeColor =8404992
                    Name ="AllocateLanes"
                    Caption ="&Allocate Lanes"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    TextFontFamily =34
                    Left =9000
                    Top =432
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    ForeColor =8404992
                    Name ="Edit"
                    Caption ="&Edit Team"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =135
                    Top =120
                    Width =8655
                    Height =6690
                    TabIndex =8
                    Name ="TabCtl19"
                    FontName ="Tahoma"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =270
                            Top =525
                            Width =8385
                            Height =6150
                            Name ="Page20"
                            Caption =" Team Details"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =223
                                    TextAlign =2
                                    Left =4541
                                    Top =612
                                    Width =1320
                                    Height =255
                                    FontWeight =700
                                    Name ="Text2"
                                    Caption ="Code"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =702
                                    Top =612
                                    Width =2505
                                    Height =270
                                    FontWeight =700
                                    Name ="Text3"
                                    Caption ="Description"
                                    FontName ="Tahoma"
                                End
                                Begin Subform
                                    OverlapFlags =215
                                    OldBorderStyle =0
                                    SpecialEffect =3
                                    Left =362
                                    Top =883
                                    Width =8292
                                    Height =5774
                                    Name ="HS SF"
                                    SourceObject ="Form.House Summary SF"
                                    EventProcPrefix ="HS_SF"

                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =223
                                    TextAlign =2
                                    Left =6153
                                    Top =600
                                    Width =1275
                                    Height =285
                                    FontWeight =700
                                    Name ="Text10"
                                    Caption ="Comp. Pool"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    TextAlign =2
                                    Left =7720
                                    Top =612
                                    Width =465
                                    Height =240
                                    FontWeight =700
                                    Name ="Text11"
                                    Caption ="Inc."
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =270
                            Top =525
                            Width =8385
                            Height =6150
                            Name ="Page21"
                            Caption ="Instructions"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =481
                                    Top =651
                                    Width =990
                                    Height =210
                                    FontWeight =700
                                    Name ="Label22"
                                    Caption ="General"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =680
                                    Top =907
                                    Width =7350
                                    Height =360
                                    FontWeight =300
                                    Name ="Label24"
                                    Caption ="Add each team that will be competing in this carnival by pushing the 'Add Team' "
                                        "button."
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =480
                                    Top =1350
                                    Width =1860
                                    Height =210
                                    FontWeight =700
                                    Name ="Label26"
                                    Caption ="Allocate Extra Points"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =675
                                    Top =1605
                                    Width =7350
                                    Height =600
                                    FontWeight =300
                                    Name ="Label27"
                                    Caption ="If you wish to give a team extra points for, say cheerleading, do so by pushing "
                                        "the 'Allocate Extra Points' button.  You can also deduct points from team this w"
                                        "ay.  Simply enter a negative value for the points."
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =465
                                    Top =2340
                                    Width =1860
                                    Height =210
                                    FontWeight =700
                                    Name ="Label28"
                                    Caption ="Allocate Lanes"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =660
                                    Top =2595
                                    Width =7350
                                    Height =600
                                    FontWeight =300
                                    Name ="Label29"
                                    Caption ="When adding competitors to an event the program will automatically place them in"
                                        " the lanes you specify in the 'Allocate Default lanes' form.  You can specify an"
                                        "y number lanes for a team.  Puch the 'Allocate Lanes' button."
                                    FontName ="Tahoma"
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

Option Compare Database   'Use database order for string comparisons

Private Sub AddB_Click()
On Error GoTo Err_AddB_Click

    OpenFormType = "ADD"
    DoCmd.OpenForm "Houses", , , , , acDialog
    Me![HS SF].Requery

    OpenFormType = "EDIT"

Exit_AddB_Click:
    Exit Sub

Err_AddB_Click:
    MsgBox Error$
    Resume Exit_AddB_Click
    
End Sub

Private Sub AllocateLanes_Click()

    DoCmd.OpenForm "Lanes", , , , , acDialog

End Sub

Private Sub Button7_Click()

End Sub

Private Sub Close_Click()
On Error GoTo Err_Close_Click


    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

Private Sub DeleteB_Click()

    Hid = Me![HS SF].Form![H_ID]
    If IsNull(Hid) Then
        MsgBox ("You must select a house or school in the list.")
    Else
        WarningMessage = "If you proceed with this delete action all information related to the team (its competitors AND ALL records set by this team) will be lost.  Are you sure you want to continue?"
        Response = MsgBox(WarningMessage, vbYesNo + vbCritical + vbDefaultButton2, "Confirm Team Delete")
        
        If Response = vbYes Then 'Yes
            Q = "DELETE DISTINCTROW [House Points-Extra].H_ID FROM [House Points-Extra]"
            Q = Q & " WHERE [House Points-Extra].H_ID=" & Hid
            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
            
            Q = "DELETE DISTINCTROW House.H_ID FROM House"
            Q = Q & " WHERE House.H_ID=" & Hid
            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
            Me![HS SF].Requery
        End If
        
        
    End If


End Sub

Private Sub Edit_Click()

    If Not IsNull(Me![HS SF].Form![H_ID]) Then
        DoCmd.OpenForm "Houses", , , "[H_ID]=" & Me![HS SF].Form![H_ID], , acDialog
    End If

End Sub

Private Sub Extra_Click()

      DoCmd.OpenForm "House Points-Extra", , , , , acDialog

End Sub

Private Sub Form_Load()

    OpenFormType = "EDIT"

End Sub

Private Sub Summary_DblClick(Cancel As Integer)

    'If Not IsNull([Summary]) Then
    '    OpenFormType = "Edit"
    '
    '    DoCmd OpenForm "Houses", , , "[H_ID]=" & [Summary], , acDialog
    '
    '    [Summary].Requery
    '
    'End If

End Sub
