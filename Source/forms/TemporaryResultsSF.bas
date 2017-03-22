Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =25
    GridY =25
    Width =4266
    ItemSuffix =6
    Left =5955
    Top =1080
    Right =11520
    Bottom =4515
    RecSrcDt = Begin
        0x72946ef6abcde140
    End
    RecordSource ="SELECT DISTINCTROW [Temporary Results-Place Order].Place, [Temporary Results-Pla"
        "ce Order].Lane, [Temporary Results-Place Order].Results FROM [Temporary Results-"
        "Place Order] ORDER BY [Temporary Results-Place Order].Place;"
    BeforeUpdate ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =288
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Width =576
                    Name ="Place"
                    ControlSource ="Place"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1498
                    Width =1131
                    TabIndex =2
                    Name ="Results"
                    ControlSource ="Results"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ListWidth =820
                    Left =634
                    Width =805
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Field4"
                    ControlSource ="Lane"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Temporary Results-Place Order].AvailableLanes FROM [Temporary Results-Pl"
                        "ace Order] ORDER BY [Temporary Results-Place Order].AvailableLanes;"
                    ColumnWidths ="570"

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

Private Sub Form_BeforeUpdate(Cancel As Integer)

    If IsNull([Lane]) And Not IsNull([Results]) Then
        MsgBox ("You must enter a lane.")
        Cancel = True
    End If
    
End Sub
