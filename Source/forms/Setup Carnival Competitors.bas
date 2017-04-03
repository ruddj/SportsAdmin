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
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =10540
    ItemSuffix =20
    Left =3255
    Top =2745
    Right =13800
    Bottom =5565
    RecSrcDt = Begin
        0x1bf2bb290fcde140
    End
    Caption ="Setup Carnival Competitors"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Rectangle
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
        Begin Section
            Height =2834
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =226
                    Top =1270
                    Width =3135
                    Height =510
                    FontSize =8
                    FontWeight =400
                    ForeColor =255
                    Name ="But2"
                    Caption ="Enter Carnival Competitor Manually"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =113
                    Top =226
                    Width =10315
                    Height =1690
                    Name ="Box46"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    Left =3628
                    Top =340
                    Width =6510
                    Height =645
                    Name ="Text0"
                    Caption ="If you can generate a plain text file with competitor information from an admini"
                        "stration system in the format given in the Help file, you can import it using th"
                        "is option.  "
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =226
                    Top =396
                    Width =3135
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    ForeColor =255
                    Name ="But1"
                    Caption ="Import Carnival Competitors from a Text File"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    Left =3630
                    Top =1365
                    Width =6510
                    Height =390
                    Name ="Text6"
                    Caption ="Competitor information can be entered manually using this option."
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9240
                    Top =2154
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =1365
                    Top =960
                    Width =765
                    Height =225
                    FontWeight =700
                    Name ="Text8"
                    Caption ="AND / OR"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =226
                    Top =2097
                    Width =3135
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =8388736
                    Name ="but3"
                    Caption ="Enter Competitors into Events"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =3630
                    Top =2190
                    Width =5220
                    Height =435
                    Name ="Text19"
                    Caption ="Competitors need to be added to each event.  Before doing this ensure that the h"
                        "ouse or school default lane allocation has been setup correctly."
                    FontName ="Tahoma"
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

Private Sub But1_Click()

On Error GoTo Err_But1_Click

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "Import Competitors"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog

Exit_But1_Click:
    Exit Sub

Err_But1_Click:
    MsgBox Error$
    Resume Exit_But1_Click
    
End Sub

Private Sub But2_Click()

On Error GoTo Err_But2_Click

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "CompetitorsSummary"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog

Exit_But2_Click:
    Exit Sub

Err_But2_Click:
    MsgBox Error$
    Resume Exit_But2_Click

End Sub

Private Sub but3_Click()

On Error GoTo Err_But3_Click

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "CompEventsSummary"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog

Exit_But3_Click:
    Exit Sub

Err_But3_Click:
    MsgBox Error$
    Resume Exit_But3_Click

End Sub

Private Sub Close_Click()

    DoCmd.Close

End Sub
