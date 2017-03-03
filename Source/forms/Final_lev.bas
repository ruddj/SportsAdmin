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
    DividingLines = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =20
    Width =9309
    ItemSuffix =30
    Left =4170
    Top =2820
    Right =13485
    Bottom =7560
    HelpContextId =530
    Filter ="[ET_Code] = Forms![EventType]![ET_Code]"
    RecSrcDt = Begin
        0xcc9a6c9bb6f2e140
    End
    RecordSource ="SELECT DISTINCTROW EventType.ET_Code, EventType.ET_Des FROM EventType WHERE ((Ev"
        "entType.ET_Code=[Forms]![EventType]![ET_Code]));"
    Caption ="Maintain Heats and Finals"
    HelpFile ="sports.hlp"
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
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =0
            BackColor =128
            Name ="FormHeader1"
        End
        Begin Section
            Height =4762
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =144
                    Top =639
                    Width =7652
                    Height =3621
                    Name ="Final_Lev_Sub"
                    SourceObject ="Form.Final_Lev_Sub"
                    LinkChildFields ="ET_Code"
                    LinkMasterFields ="ET_Code"
                    OnExit ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =7965
                    Top =75
                    Width =1284
                    Height =705
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Create"
                    Caption ="&Create Heats and Finals"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =170
                    Top =113
                    Width =540
                    Height =270
                    FontWeight =400
                    Name ="Text20"
                    Caption ="Event:"
                    FontName ="Arial"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    Left =807
                    Top =113
                    Width =6996
                    Height =283
                    FontWeight =700
                    TabIndex =2
                    BackColor =12632256
                    Name ="Field21"
                    ControlSource ="ET_Des"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =7965
                    Top =4170
                    Width =1284
                    Height =510
                    FontSize =8
                    TabIndex =3
                    Name ="Button17"
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
                    Left =7965
                    Top =2914
                    Width =1284
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    HelpContextId =530
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =8037
                    Top =1083
                    Width =1011
                    TabIndex =5
                    Name ="ET_Code"
                    ControlSource ="ET_Code"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2616
                    Top =4453
                    TabIndex =6
                    Name ="ClearExisting"
                    DefaultValue ="True"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =4425
                            Width =2355
                            Height =240
                            FontWeight =400
                            Name ="Label29"
                            Caption ="Remove all existing heats"
                        End
                    End
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

Private Sub Button17_Click()
On Error GoTo Err_Button17_Click


    DoCmd.Close

Exit_Button17_Click:
    Exit Sub

Err_Button17_Click:
    MsgBox Error$
    Resume Exit_Button17_Click
    
End Sub


Private Sub Create_Click()
On Error GoTo Err_Create_Click

  If AutomaticallyCreateHeatsAndFinals(Me!ET_Code, , , Me!ClearExisting) Then
    Response = MsgBox("Heats and finals have been successfully created.", vbInformation)
    DoCmd.Close acForm, Me.Name
  End If
  
Exit_Create_Click:

  DoCmd.RunMacro "ClosePleaseWait"
  Exit Sub


Err_Create_Click:

    MsgBox Error$
    Resume Exit_Create_Click

End Sub

Private Sub Final_Lev_Sub_Exit(Cancel As Integer)
    
    If Not CheckFinalIntegrity(Me![ET_Code], "Events") Then
         Response = MsgBox("Finals should be in consecutive order starting at 0 and increasing by one (1) only.  This is not necessary but is recommended.  Do you wish to continue?", 20, "Final Integrity Warning")
         If Response = 6 Then
            Cancel = False
         Else
            Cancel = True
         End If
    End If

End Sub
