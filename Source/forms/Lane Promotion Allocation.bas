Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    GridX =20
    GridY =20
    Width =3636
    ItemSuffix =33
    Left =3735
    Top =510
    Right =11985
    Bottom =8520
    HelpContextId =520
    RecSrcDt = Begin
        0x2af6e5b911cde140
    End
    Caption ="Lane Promotion Allocation"
    HelpFile ="SportsAdmin.chm"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin OptionButton
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =187
            Height =187
        End
        Begin OptionGroup
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin ToggleButton
            TextFontFamily =2
            Width =283
            Height =283
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4701
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =215
                    SpecialEffect =3
                    Left =151
                    Top =680
                    Width =1938
                    Height =3863
                    Name ="LPA"
                    SourceObject ="Form.Lane Promotion Allocation SF"
                    LinkChildFields ="ET_Code"
                    LinkMasterFields ="ET_Code"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =453
                    Top =453
                    Width =630
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text10"
                    Caption ="Place"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =1028
                    Top =453
                    Width =675
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text12"
                    Caption ="Lane"
                    FontName ="Arial"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =2474
                    Top =2979
                    Width =786
                    TabIndex =1
                    Name ="ET_Code"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =2417
                    Top =4056
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="Close But"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Close_But"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =2304
                    Top =504
                    Width =1318
                    Height =1633
                    TabIndex =3
                    Name ="Field26"
                    DefaultValue ="1"

                    Begin
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontFamily =34
                            Left =2417
                            Top =617
                            Width =1134
                            Height =510
                            FontSize =8
                            FontWeight =400
                            OptionValue =1
                            Name ="Place"
                            Caption ="Order by Place"
                            FontName ="Arial"
                            OnMouseDown ="[Event Procedure]"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            TextFontFamily =34
                            Left =2417
                            Top =1241
                            Width =1134
                            Height =510
                            FontSize =8
                            FontWeight =400
                            OptionValue =2
                            Name ="Lane"
                            Caption ="Order by  Lane"
                            FontName ="Arial"
                            OnMouseDown ="[Event Procedure]"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    Left =72
                    Top =72
                    Width =3456
                    Height =283
                    FontWeight =700
                    TabIndex =4
                    BackColor =12632256
                    Name ="Field21"
                    ControlSource ="=[Forms]![EventType]![ET_Des]"

                End
                Begin Line
                    OverlapFlags =87
                    Left =2232
                    Top =360
                    Width =0
                    Height =4320
                    Name ="Line32"
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

Private Sub Close_But_Click()
On Error GoTo Err_Close_But_Click

    x = DCount("[ET_Code]", "Lane Promotion Allocation", "[ET_Code]=" & Me![ET_Code])
    
    If x < Forms![EventType]![Lane_Cnt] Then
        Response = MsgBox("The number of lanes you have set up is less than the Lane / Competitor count specified on the previous form.  Do you still wish to continue?", 36, "Too few lanes?")
        If Response = 6 Then
            DoCmd.Close
        End If
        
    Else
        DoCmd.Close
    End If

Exit_Close_But_Click:
    Exit Sub

Err_Close_But_Click:
    MsgBox Error$
    Resume Exit_Close_But_Click
    
End Sub

Private Sub Form_Load()

    [ET_Code].Value = Forms![EventType]![ET_Code]

End Sub

Private Sub Lane_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Q = "SELECT DISTINCTROW [Lane Promotion Allocation].ET_Code, [Lane Promotion Allocation].Place, [Lane Promotion Allocation].Lane "
    Q = Q & "FROM [Lane Promotion Allocation] "
    Q = Q & "ORDER BY [Lane Promotion Allocation].Lane"

    Me![LPA].Form.RecordSource = Q
    Me![LPA].Form.Requery
    

End Sub

Private Sub Place_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Q = "SELECT DISTINCTROW [Lane Promotion Allocation].ET_Code, [Lane Promotion Allocation].Place, [Lane Promotion Allocation].Lane "
    Q = Q & "FROM [Lane Promotion Allocation] "
    Q = Q & "ORDER BY [Lane Promotion Allocation].Place"

    Me![LPA].Form.RecordSource = Q
    Me![LPA].Form.Requery

End Sub
