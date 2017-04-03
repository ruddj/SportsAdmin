Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =5256
    ItemSuffix =20
    Left =4170
    Top =3240
    Right =7980
    Bottom =8010
    HelpContextId =140
    RecSrcDt = Begin
        0x7903b8b911cde140
    End
    RecordSource ="SELECT DISTINCTROW UCase([H_Code]) AS Hcode, House.H_NAme, House.Flag FROM House"
        " WHERE ((House.Include=True)) ORDER BY UCase([H_Code]);"
    Caption ="ExportTextSF1"
    HelpFile ="SportsAdmin.chm"
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
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =12632256
        End
        Begin ListBox
            AutoLabel = NotDefault
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =12632256
        End
        Begin ComboBox
            AutoLabel = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            BackColor =12632256
        End
        Begin PageBreak
            Width =283
        End
        Begin FormHeader
            Height =225
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =72
                    Width =525
                    Height =210
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Label18"
                    Caption ="Code"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1152
                    Width =525
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Label19"
                    Caption ="Name"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            Height =288
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =56
                    Width =1020
                    Height =225
                    BackColor =16777215
                    Name ="H_Code"
                    ControlSource ="Hcode"
                    StatusBarText ="House / School Code ie. Asher, COC, Beaudesert, Australia, Individual?)"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =1148
                    Width =2100
                    Height =225
                    TabIndex =1
                    BackColor =16777215
                    Name ="H_NAme"
                    ControlSource ="H_NAme"
                    StatusBarText ="House / School Name"
                    FontName ="Tahoma"

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =3344
                    TabIndex =2
                    Name ="Flag"
                    ControlSource ="Flag"
                    StatusBarText ="Genreal Flag"
                    ControlTipText ="Tick the teams you want to include."

                End
            End
        End
        Begin FormFooter
            Height =396
            BackColor =-2147483633
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =2779
                    Top =56
                    Width =615
                    Height =225
                    FontSize =6
                    FontWeight =400
                    Name ="SelectAllBut"
                    Caption ="ALL"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Select all teams."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =1870
                    Top =56
                    Width =615
                    Height =225
                    FontSize =6
                    FontWeight =400
                    TabIndex =1
                    Name ="DEselectBut"
                    Caption ="NONE"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Select no teams."

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

Private Sub Button17_Click()

End Sub

Private Sub DEselectBut_Click()

On Error GoTo Err_DEselectBut_Click

    Q = "UPDATE DISTINCTROW House SET House.Flag = No"
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
                          
    Me.Refresh

Exit_DEselectBut_Click:
    Exit Sub

Err_DEselectBut_Click:
    MsgBox Error$
    Resume Exit_DEselectBut_Click

End Sub

Private Sub SelectAllBut_Click()
On Error GoTo Err_SelectAllBut_Click

    Q = "UPDATE DISTINCTROW House SET House.Flag = Yes"
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
                          
    Me.Refresh

Exit_SelectAllBut_Click:
    Exit Sub

Err_SelectAllBut_Click:
    MsgBox Error$
    Resume Exit_SelectAllBut_Click
    
End Sub
