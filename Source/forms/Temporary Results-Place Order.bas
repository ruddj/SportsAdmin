Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    GridX =25
    GridY =25
    Width =3398
    ItemSuffix =7
    Left =1380
    Top =540
    Right =11520
    Bottom =7590
    RecSrcDt = Begin
        0x72946ef6abcde140
    End
    Caption ="Enter Results in Place Order"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =3798
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =3
                    Left =113
                    Top =343
                    Width =3207
                    Height =2718
                    Name ="Embedded0"
                    SourceObject ="Form.TemporaryResultsSF"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =396
                    Top =113
                    Width =480
                    Height =210
                    Name ="Text2"
                    Caption ="Place"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =1077
                    Top =116
                    Width =480
                    Height =210
                    Name ="Text3"
                    Caption ="Lane"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =2040
                    Top =113
                    Width =525
                    Height =210
                    Name ="Text4"
                    Caption ="Result"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1984
                    Top =3174
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Close"
                    Caption ="Enter Results"
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
                    Left =113
                    Top =3174
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="Cancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

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

    GlobalCancel = False
    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub
