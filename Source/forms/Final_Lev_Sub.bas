Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =7125
    ItemSuffix =32
    Left =2160
    Top =1275
    Right =10950
    Bottom =5055
    RecSrcDt = Begin
        0x1c945709704ae240
    End
    RecordSource ="SELECT DISTINCTROW Final_Lev.ET_Code, Final_Lev.F_Lev, Final_Lev.NoHeats, Final_"
        "Lev.PtScale, Final_Lev.ProType, Final_Lev.UseTimes, Final_Lev.ProNum, Final_Lev."
        "EffectsRecords FROM Final_Lev ORDER BY Final_Lev.F_Lev;"
    BeforeUpdate ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0xa2050000a1050000a1050000a105000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            AutoLabel = NotDefault
            OldBorderStyle =0
            LabelAlign =3
            TextAlign =2
            BorderLineStyle =0
            Width =555
            Height =230
            LabelX =-236
            FontName ="Arial"
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
            Height =455
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =113
                    Width =540
                    Height =450
                    FontWeight =400
                    Name ="Text13"
                    Caption ="Final Level"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =750
                    Top =5
                    Width =574
                    Height =450
                    FontWeight =400
                    Name ="Text15"
                    Caption ="# of Heats"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1398
                    Top =114
                    Width =1069
                    Height =240
                    FontWeight =400
                    Name ="Text17"
                    Caption ="Point Scale"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2694
                    Top =2
                    Width =964
                    Height =450
                    FontWeight =400
                    Name ="Text19"
                    Caption ="Promotion Method"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =3825
                    Width =1200
                    Height =450
                    FontWeight =400
                    Name ="Text21"
                    Caption ="Use Results (X) / Places ()"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =6225
                    Width =900
                    Height =450
                    FontWeight =400
                    Name ="Text28"
                    Caption ="# of Comp. in Heat"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    Left =5025
                    Width =1200
                    Height =450
                    FontWeight =400
                    Name ="Label30"
                    Caption ="Results effect event records."
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =307
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =750
                    Width =570
                    Height =256
                    TabIndex =1
                    Name ="NoHeats"
                    ControlSource ="NoHeats"
                    StatusBarText ="The number of heats in this final level."
                    ValidationRule =">0 And <1000"
                    ValidationText ="You must enter a number between 0 and 1000 (or push Esc to cancel)."
                    ControlTipText ="The number of heats in this final level."

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =4335
                    Width =232
                    Height =275
                    TabIndex =4
                    Name ="UseTimes"
                    ControlSource ="UseTimes"
                    StatusBarText ="Use either Time or Placing results to select competitors to promote to the next "
                        "final level"
                    DefaultValue ="Yes"
                    ControlTipText ="Tick if results are used to determine which competitors make it into the next fi"
                        "nal."

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    ListWidth =1225
                    Left =1395
                    Width =1285
                    Height =256
                    TabIndex =2
                    ColumnInfo ="\"\";\">\";\"10\";\"20\""
                    Name ="PtScale"
                    ControlSource ="PtScale"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT PointsScale.PtScale FROM PointsScale;"
                    ColumnWidths ="975"
                    FontName ="Arial"
                    ControlTipText ="The pointscale used to allocate points for the final level."

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ColumnCount =2
                    ListWidth =2800
                    Left =2766
                    Width =1120
                    Height =256
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="ProType"
                    ControlSource ="ProType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Promotion.ProType, Promotion.Desc FROM Promotion;"
                    ColumnWidths ="900;1650"
                    FontName ="Arial"
                    ControlTipText ="The manner competitiors are promoted into the next final level."

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =6279
                    Width =690
                    Height =256
                    TabIndex =6
                    Name ="Field29"
                    ControlSource ="ProNum"
                    StatusBarText ="Optional: This value is used to promote competitors in events that don't use lan"
                        "es."
                    ControlTipText ="Optional: This value is used to promote competitors in events that don't use lan"
                        "es."

                End
                Begin ComboBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    DecimalPlaces =0
                    ColumnCount =2
                    ListWidth =2268
                    Left =68
                    Width =601
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"2\";\"1\""
                    Name ="F_Lev"
                    ControlSource ="F_Lev"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Final Level Sub].F_Lev, [Final Level Sub].F_Lev_Sub FROM [Final Level Su"
                        "b] ORDER BY [Final Level Sub].F_Lev;"
                    ColumnWidths ="567;1701"
                    StatusBarText ="Final Levels start at 0 (0: Grand Final, 1: Semi-Final, 2: Quarter-Final etc)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Final Levels start at 0 (0: Grand Final, 1: Semi-Final, 2: Quarter-Final etc)"

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =5535
                    Width =232
                    Height =275
                    TabIndex =5
                    Name ="EffectsRecords"
                    ControlSource ="EffectsRecords"
                    StatusBarText ="Use either Time or Placing results to select competitors to promote to the next "
                        "final level"
                    DefaultValue ="Yes"
                    ControlTipText ="Tick if results effect the records for this event."

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

Private Sub F_Lev_AfterUpdate()

  If Me!F_Lev = 0 Then Me!ProType = "NONE"
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    
    If IsNull(Me![ProType]) Then
        Response = MsgBox("You must enter a valid promotion type (or push Esc to cancel).", vbInformation)
        Cancel = True
    ElseIf IsNull(Me![PtScale]) Then
        Response = MsgBox("You must enter a valid Point Scale (or push Esc to cancel).", vbInformation)
        Cancel = True
    ElseIf IsNull(Me![NoHeats]) Then
        Response = MsgBox("You must enter the number of heats (or push Esc to cancel).", vbInformation)
        Cancel = True
    End If

End Sub
