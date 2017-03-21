Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridY =10
    Width =9490
    ItemSuffix =41
    Left =2400
    Top =510
    Right =11520
    Bottom =7440
    HelpContextId =40
    RecSrcDt = Begin
        0x119d78290fcde140
    End
    RecordSource ="SELECT DISTINCTROW House.H_Code, House.H_NAme, House.HT_Code, House.Include, Hou"
        "se.Details, House.Lane, House.CompPool, House.Flag, House.H_ID FROM House ORDER "
        "BY House.H_NAme;"
    Caption ="House"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
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
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            Height =324
            BackColor =-2147483633
            Name ="Detail0"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3741
                    Width =1710
                    Height =285
                    TabIndex =1
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    Format =">"
                    StatusBarText ="House / School Code ie. Asher, COC, Beaudesert, Australia, Individual?)"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Enter a short code for the team."

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =60
                    Width =3630
                    Height =285
                    Name ="H_NAme"
                    ControlSource ="H_NAme"
                    StatusBarText ="House / School Name"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Enter a descriptive name for the team."

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =7140
                    Top =60
                    Height =217
                    ColumnWidth =960
                    TabIndex =3
                    Name ="Include"
                    ControlSource ="Include"
                    StatusBarText ="Include house in carnival"
                    ControlTipText ="Tick if you want to include this team in the carnival."

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5490
                    Width =1320
                    Height =285
                    TabIndex =2
                    Name ="CompPool"
                    ControlSource ="CompPool"
                    Format ="General Number"
                    StatusBarText ="The number of people the House / School has available to them"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="The size of the school or region that the team was selected from."

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =8485
                    Width =1005
                    TabIndex =4
                    Name ="H_ID"
                    ControlSource ="H_ID"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =7464
                    Width =1005
                    Height =285
                    TabIndex =5
                    Name ="Field32"
                    ControlSource ="HT_Code"

                End
            End
        End
        Begin FormFooter
            Height =510
            BackColor =-2147483633
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5102
                    Top =56
                    Width =1134
                    Height =397
                    FontSize =8
                    FontWeight =400
                    Name ="Select All"
                    Caption ="Include All"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Select_All"
                    ControlTipText ="Tick the 'Include' box for all teams."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3855
                    Top =56
                    Width =1134
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Exclude All"
                    Caption ="Exclude All"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Exclude_All"
                    ControlTipText ="Untick the 'Include' box for all competitors."

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

Private Sub Button35_Click()

End Sub

Private Sub CompPool_DblClick(Cancel As Integer)

    EditHouse

End Sub

Private Sub Detail0_DblClick(Cancel As Integer)

    EditHouse

End Sub

Private Sub EditHouse()

    DoCmd.OpenForm "Houses", , , "[H_ID]=" & Me![H_ID], , acDialog

End Sub

Private Sub Exclude_All_Click()

    Q = "UPDATE DISTINCTROW House SET House.Include = No "
    'q = q & "WHERE House.HT_Code like " & Forms![House SUmmary].Form![HT_Code]

    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
    
    Me.Refresh

End Sub

Private Sub H_Code_DblClick(Cancel As Integer)

    EditHouse

End Sub

Private Sub H_NAme_DblClick(Cancel As Integer)

    EditHouse

End Sub

Private Sub Select_All_Click()

    Q = "UPDATE DISTINCTROW House SET House.Include = Yes "
    'q = q & "WHERE House.HT_Code like " & Forms![House SUmmary].Form![HT_Code]

    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
    
    Me.Refresh
    
End Sub
