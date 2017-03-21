Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    GridX =20
    GridY =20
    Width =8965
    ItemSuffix =12
    Left =-17865
    Top =4365
    Right =-7575
    Bottom =10725
    HelpContextId =40
    RecSrcDt = Begin
        0x403f3e042dc7e140
    End
    Caption ="Special Points Allocation"
    OnCurrent ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
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
            Height =4025
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =7665
                    Top =3510
                    Width =1059
                    Height =435
                    FontSize =8
                    FontWeight =400
                    Name ="Close"
                    Caption ="&Done"
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
                    Left =120
                    Top =3510
                    Width =1059
                    Height =435
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    HelpContextId =40
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =120
                    Top =75
                    Width =8625
                    Height =3405
                    TabIndex =2
                    Name ="TabCtl8"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =255
                            Top =480
                            Width =8355
                            Height =2865
                            Name ="Page9"
                            Caption ="Extra Points"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin ListBox
                                    OverlapFlags =215
                                    ColumnCount =2
                                    Left =346
                                    Top =716
                                    Width =1810
                                    Height =2561
                                    BorderColor =12632256
                                    Name ="H_ID"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT House.H_ID, House.H_NAme, House.Include FROM House WHERE ((House.Include="
                                        "Yes));"
                                    ColumnWidths ="0;1562"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Click on the team you wish to add points to."

                                End
                                Begin Subform
                                    OverlapFlags =215
                                    SpecialEffect =3
                                    Left =2387
                                    Top =716
                                    Width =6223
                                    Height =2538
                                    Name ="SF"
                                    SourceObject ="Form.House Points-Extra SF"
                                    LinkChildFields ="H_ID"
                                    LinkMasterFields ="H_ID"

                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =345
                                    Top =495
                                    Width =1200
                                    Height =225
                                    FontWeight =700
                                    Name ="Text4"
                                    Caption ="Teams"
                                    FontName ="Arial"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =255
                            Top =480
                            Width =8355
                            Height =2865
                            Name ="Page10"
                            Caption ="Instructions"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =335
                                    Top =560
                                    Width =7597
                                    Height =2267
                                    Name ="Label11"
                                    Caption ="Click on a Team and enter the extra points you wish to allocate in the box on th"
                                        "e right.  You can enter both positive and negative points.\015\012\015\012You ca"
                                        "n also enter an explanantion as to why extra points were added or deducted."
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

Private Sub Close_Click()
On Error GoTo Err_Close_Click


    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

Private Sub Form_Current()

    If Not IsNull(Me![H_ID]) Then
        If IsNull(DLookup("[H_ID]", "House", "[H_ID]=" & Me!H_ID)) Then
            Me![SF].visible = False
        Else
            Me![SF].visible = True
        End If
    Else
        Me![SF].visible = False
    End If

        
End Sub

Private Sub H_ID_AfterUpdate()

    Form_Current

End Sub
