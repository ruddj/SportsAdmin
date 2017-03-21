Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =15953
    ItemSuffix =29
    Left =420
    Top =75
    Right =10860
    Bottom =6510
    HelpContextId =70
    RecSrcDt = Begin
        0xd2dd269aefe5e140
    End
    RecordSource ="SELECT DISTINCTROW Competitors.Include, Competitors.Surname, Competitors.Gname, "
        "Competitors.H_Code, Competitors.Age, Competitors.DOB, Competitors.Sex FROM House"
        " INNER JOIN Competitors ON House.H_Code = Competitors.H_Code WHERE (((House.Incl"
        "ude)=Yes)) ORDER BY Competitors.Surname, Competitors.Gname;"
    Caption ="Competitors-BulkMaintain SF"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            Height =288
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4537
                    Width =585
                    Height =225
                    TabIndex =3
                    Name ="Sex"
                    ControlSource ="Sex"
                    Format =">"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3528
                    Width =945
                    Height =225
                    TabIndex =2
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    Format =">"
                    StatusBarText ="House Code"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5184
                    Width =570
                    Height =225
                    TabIndex =4
                    Name ="Age"
                    ControlSource ="Age"
                    Format =">"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =72
                    Width =1725
                    Height =225
                    Name ="Gname"
                    ControlSource ="Gname"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =1872
                    Width =1575
                    Height =225
                    TabIndex =1
                    Name ="Surname"
                    ControlSource ="Surname"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5828
                    Width =960
                    Height =225
                    TabIndex =5
                    Name ="DOB"
                    ControlSource ="DOB"
                    FontName ="Arial"

                End
                Begin OptionButton
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =6962
                    TabIndex =6
                    Name ="Button26"
                    ControlSource ="Include"

                End
            End
        End
        Begin FormFooter
            Height =453
            BackColor =-2147483633
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6747
                    Top =56
                    Width =681
                    FontSize =7
                    FontWeight =400
                    Name ="All"
                    Caption ="ALL"
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
                    Left =5896
                    Top =56
                    Width =681
                    FontSize =7
                    FontWeight =400
                    TabIndex =1
                    Name ="None"
                    Caption ="None"
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

Option Explicit

Private Sub All_Click()

    Dim Q As Variant

    Q = "UPDATE DISTINCTROW Competitors SET Competitors.Include = Yes"
    'DoCmd SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
    Me.Refresh
    
End Sub

Private Sub None_Click()

    Dim Q As Variant

    Q = "UPDATE DISTINCTROW Competitors SET Competitors.Include = No"
    'DoCmd SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
    Me.Refresh

End Sub
