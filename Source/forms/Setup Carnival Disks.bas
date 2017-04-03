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
    ItemSuffix =18
    Left =3255
    Top =2850
    Right =14715
    Bottom =6855
    RecSrcDt = Begin
        0x12a5bd290fcde140
    End
    Caption ="Setup Carnival Disks"
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
            Height =2551
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =226
                    Top =1530
                    Width =3135
                    Height =510
                    FontSize =8
                    FontWeight =400
                    ForeColor =32768
                    Name ="But2"
                    Caption ="Import Carnival Diskettes"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =113
                    Top =226
                    Width =10315
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
                    Caption ="Create these disks ONLY when you have completed the setting up of ALL the events"
                        " (and ALL their respective heats)."
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
                    ForeColor =32768
                    Name ="But1"
                    Caption ="Create Carnival Diskettes"
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
                    OverlapFlags =85
                    Left =3630
                    Top =1245
                    Width =5265
                    Height =1185
                    Name ="Text6"
                    Caption ="You do not have to import all the Carnival disks at once.  They can be imported "
                        "(and then checked) as they are received.  It is advisable to check that the info"
                        "rmation has been successfully added to the database and to do a \"spot check\" o"
                        "n a number of events to ensure that the competitor information is correct."
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9240
                    Top =1870
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

    DocName = "ExportData"
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

    DocName = "Import Data"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog

Exit_But2_Click:
    Exit Sub

Err_But2_Click:
    MsgBox Error$
    Resume Exit_But2_Click

End Sub

Private Sub Close_Click()

    DoCmd.Close

End Sub
