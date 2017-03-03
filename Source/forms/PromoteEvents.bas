Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    GridY =10
    Width =5215
    ItemSuffix =12
    Left =3255
    Top =2085
    Right =15795
    Bottom =7470
    RecSrcDt = Begin
        0x2677765db6f2e140
    End
    RecordSource ="showDialog"
    Caption ="Promote Event"
    OnOpen ="[Event Procedure]"
    FilterOnLoad =0
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
        Begin Section
            Height =3118
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =566
                    Top =2551
                    Width =984
                    Height =405
                    FontSize =8
                    FontWeight =400
                    Name ="YesBut"
                    Caption ="Yes"
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
                    Left =2157
                    Top =2552
                    Width =984
                    Height =405
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="NoBut"
                    Caption ="No"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =399
                    Top =572
                    Width =4422
                    Height =1243
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483633
                    Name ="Txt"
                    FontName ="Arial"

                End
                Begin CheckBox
                    SpecialEffect =1
                    OverlapFlags =85
                    Left =1760
                    Top =2041
                    TabIndex =3
                    Name ="SD"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Left =2016
                            Top =2041
                            Width =990
                            Height =240
                            Name ="Text5"
                            Caption ="Yes to all"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =396
                    Top =226
                    Width =3300
                    Height =210
                    Name ="Text6"
                    Caption ="Are you sure you wish to promote:"
                    FontName ="Arial"
                End
                Begin CheckBox
                    Visible = NotDefault
                    SpecialEffect =1
                    OverlapFlags =85
                    Left =399
                    Top =2041
                    TabIndex =4
                    Name ="SDA"
                    ControlSource ="ShowDialog"
                    DefaultValue ="No"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3741
                    Top =2551
                    Width =984
                    Height =405
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    Name ="Cancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

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

Private Sub Cancel_Click()

    GlobalCancel = True
    DoCmd.Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    Me![Txt] = Me.OpenArgs
    
End Sub

Private Sub NoBut_Click()
On Error GoTo Err_NoBut_Click

    GlobalNo = True
    DoCmd.Close

Exit_NoBut_Click:
    Exit Sub

Err_NoBut_Click:
    MsgBox Error$
    Resume Exit_NoBut_Click
    
End Sub

Private Sub SD_AfterUpdate()

    Me![SDA] = Not (Me![SD])

End Sub

Private Sub YesBut_Click()
On Error GoTo Err_YesBut_Click

    GlobalCancel = False
    DoCmd.Close

Exit_YesBut_Click:
    Exit Sub

Err_YesBut_Click:
    MsgBox Error$
    Resume Exit_YesBut_Click
    
End Sub
