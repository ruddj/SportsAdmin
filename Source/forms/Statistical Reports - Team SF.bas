Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =4536
    ItemSuffix =21
    Left =9165
    Top =5565
    Right =11370
    Bottom =7065
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
            Height =255
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =75
                    Top =1
                    Width =525
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Label19"
                    Caption ="Name"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1469
                    Top =1
                    Width =585
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Label20"
                    Caption ="Include"
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
                    Left =72
                    Width =1575
                    Height =225
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
                    Left =1764
                    TabIndex =1
                    Name ="Flag"
                    ControlSource ="Flag"
                    StatusBarText ="Genreal Flag"
                    ControlTipText ="Tick if you want to include this event."

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
                    Left =1341
                    Top =72
                    Width =615
                    Height =270
                    FontSize =6
                    FontWeight =400
                    Name ="SelectAllBut"
                    Caption ="ALL"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Select all teams."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =432
                    Top =72
                    Width =615
                    Height =270
                    FontSize =6
                    FontWeight =400
                    TabIndex =1
                    Name ="DEselectBut"
                    Caption ="NONE"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
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

    Q = "UPDATE DISTINCTROW House SET House.Flag = Yes WHERE House.Include = True"
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
