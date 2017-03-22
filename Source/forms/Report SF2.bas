Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =12940
    ItemSuffix =29
    Left =11550
    Top =2970
    Right =13305
    Bottom =6105
    OrderBy ="[Report SF2].ET_Des"
    RecSrcDt = Begin
        0xff6a99cdede5e140
    End
    RecordSource ="SELECT EventType.ET_Code, EventType.ET_Des, EventType.Units, EventType.Lane_Cnt,"
        " EventType.R_Code, EventType.Include, EventType.EntrantNum, EventType.Flag FROM "
        "EventType WHERE (((EventType.Include)=True));"
    Caption ="EventType"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
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
        Begin TextBox
            AutoLabel = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =240
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7823
                    Top =65
                    Width =930
                    Height =105
                    FontWeight =400
                    Name ="Text8"
                    Caption ="ET_Code"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    Left =56
                    Width =825
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text12"
                    Caption ="Event"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1395
                    Width =720
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text24"
                    Caption ="Include"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =288
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7710
                    Width =1005
                    ColumnWidth =1020
                    Name ="ET_Code"
                    ControlSource ="ET_Code"
                    StatusBarText ="Event Type Code"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =72
                    Width =1590
                    TabIndex =1
                    Name ="ET_Des"
                    ControlSource ="ET_Des"
                    StatusBarText ="Event Description - ie. 200m; 100m Hurdles; High Jump- No sex / age specifics"
                    FontName ="Arial"

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =1751
                    TabIndex =2
                    Name ="Flag"
                    ControlSource ="Flag"
                    ControlTipText ="Tick to include event."

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =4365
                    Width =396
                    TabIndex =3
                    Name ="R_Code"
                    ControlSource ="R_Code"

                End
            End
        End
        Begin FormFooter
            Height =340
            BackColor =-2147483633
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1303
                    Top =56
                    Width =570
                    Height =255
                    FontSize =7
                    FontWeight =400
                    Name ="All"
                    Caption ="ALL"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Include all events."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =113
                    Top =56
                    Width =630
                    Height =255
                    FontSize =7
                    FontWeight =400
                    TabIndex =1
                    Name ="None"
                    Caption ="NONE"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Include no events."

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

Private Sub All_Click()
On Error GoTo Err_All_Click

    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE DISTINCTROW EventType SET EventType.Flag = True WHERE EventType.Include = True"
    DoCmd.SetWarnings True

    Me.Refresh

Exit_All_Click:
    Exit Sub

Err_All_Click:
    MsgBox Error$
    Resume Exit_All_Click
    
End Sub

Private Sub Button26_Click()
On Error GoTo Err_Button26_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Button26_Click:
    Exit Sub

Err_Button26_Click:
    MsgBox Error$
    Resume Exit_Button26_Click
    
End Sub

Private Sub None_Click()
On Error GoTo Err_None_Click


    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE DISTINCTROW EventType SET EventType.Flag = False"
    DoCmd.SetWarnings True

    Me.Refresh

Exit_None_Click:
    Exit Sub

Err_None_Click:
    MsgBox Error$
    Resume Exit_None_Click
    
End Sub
