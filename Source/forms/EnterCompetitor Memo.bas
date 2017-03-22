Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridY =10
    Width =5102
    ItemSuffix =8
    Left =3915
    Top =450
    Right =11520
    Bottom =6480
    RecSrcDt = Begin
        0x65926a5db6f2e140
    End
    RecordSource ="Temporary Memo"
    Caption ="Memo Attached to Event"
    HelpFile ="SportsAdmin.chm"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
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
        Begin Section
            Height =4138
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    Left =226
                    Top =920
                    Width =4649
                    Height =2479
                    Name ="Field0"
                    ControlSource ="Memo"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3741
                    Top =3514
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Close"
                    Caption ="Update"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =226
                    Top =3514
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="Cancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =225
                    Top =120
                    Width =3810
                    Height =210
                    Name ="Text4"
                    Caption ="Enter the information you require for the event below."
                End
                Begin ComboBox
                    OverlapFlags =85
                    Left =1867
                    Top =510
                    Width =2271
                    TabIndex =3
                    Name ="Competitor"
                    RowSourceType ="Table/Query"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =850
                            Top =514
                            Width =870
                            Height =240
                            Name ="Text6"
                            Caption ="Competitor:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4251
                    Top =510
                    Width =336
                    Height =291
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    Name ="Add"
                    Caption ="Add"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadada4adadadaddadada444adadadaadada44444adadaddada4444444adada ,
                        0xada444444444adadda44444444444adaadadad444dadadaddadada444adadada ,
                        0xadadad444dadadaddadada444adadadaadadad444dadadaddadada444adadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

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

Private Sub Add_Click()
On Error GoTo Err_Add_Click

    Me![Memo] = Me![Memo] & "     " & [Competitor].Value     ' Chr$(13) & Chr$(10)
    Me.Refresh

Exit_Add_Click:
    Exit Sub

Err_Add_Click:
    MsgBox Error$
    Resume Exit_Add_Click
    
End Sub

Private Sub Cancel_Click()
On Error GoTo Err_Cancel_Click

    GlobalCancel = True
    DoCmd.Close

Exit_Cancel_Click:
    Exit Sub

Err_Cancel_Click:
    MsgBox Error$
    Resume Exit_Cancel_Click
    
End Sub

Private Sub Close_Click()
On Error GoTo Err_Close_Click

    DoCmd.RunCommand acCmdSaveRecord
    Me![Memo] = Trim(Me![Memo])

    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

Private Sub Competitor_DblClick(Cancel As Integer)

    Add_Click

End Sub

Private Sub Form_Load()


    Me![Competitor].RowSource = Me.OpenArgs

End Sub
