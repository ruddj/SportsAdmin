Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    GridX =20
    GridY =20
    Width =10956
    ItemSuffix =113
    Left =570
    Top =240
    Right =14865
    Bottom =9600
    HelpContextId =270
    RecSrcDt = Begin
        0x49d6923c4fcce140
    End
    RecordSource ="Misc-Statistics"
    Caption ="Generate Statistical Reports"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    RibbonName ="SportsMenu"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =6689
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =3060
                    Top =150
                    Width =3027
                    Height =3746
                    Name ="Box53"
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =150
                    Top =4005
                    Width =2862
                    Height =2576
                    Name ="Box54"
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =144
                    Top =144
                    Width =2847
                    Height =3746
                    Name ="Box55"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =9600
                    Top =6101
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =6408
                    Top =432
                    Width =2633
                    Height =3771
                    TabIndex =1
                    Name ="EventSF"
                    SourceObject ="Form.Report SF2"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6264
                    Top =144
                    Width =1515
                    Height =225
                    Name ="Text9"
                    Caption ="Selected events "
                    FontName ="Tahoma"
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =6405
                    Top =4713
                    Width =2640
                    Height =1874
                    TabIndex =2
                    Name ="Embedded0"
                    SourceObject ="Form.Statistical Reports - Team SF"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6405
                    Top =4425
                    Width =2115
                    Height =225
                    Name ="Text10"
                    Caption ="Selected Teams:"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9507
                    Top =150
                    Width =1239
                    Height =510
                    FontSize =7
                    FontWeight =400
                    TabIndex =3
                    Name ="Selected"
                    Caption ="Preview Selected Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    SpecialEffect =1
                    OverlapFlags =215
                    Left =427
                    Top =370
                    TabIndex =4
                    Name ="Overall"
                    ControlSource ="Overall"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =680
                            Top =314
                            Width =1785
                            Height =300
                            ForeColor =8388608
                            Name ="Text33"
                            Caption ="Overall Results^"
                            FontName ="Tahoma"
                            LayoutCachedLeft =680
                            LayoutCachedTop =314
                            LayoutCachedWidth =2465
                            LayoutCachedHeight =614
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =427
                    Top =790
                    TabIndex =5
                    Name ="byAge"
                    ControlSource ="Age"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =674
                            Top =731
                            Width =1815
                            Height =300
                            ForeColor =8388608
                            Name ="Text36"
                            Caption ="Overall Results by Age^"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =427
                    Top =1210
                    TabIndex =6
                    Name ="bySex"
                    ControlSource ="Gender"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =674
                            Top =1151
                            Width =2055
                            Height =300
                            ForeColor =8388608
                            Name ="Text38"
                            Caption ="Overall Results by Gender^"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =427
                    Top =1630
                    TabIndex =7
                    Name ="bySexAge"
                    ControlSource ="Gender-Age"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =674
                            Top =1571
                            Width =2190
                            Height =300
                            ForeColor =8388608
                            Name ="Text40"
                            Caption ="Overall Res. by Gender/Age^"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =223
                    Left =495
                    Top =4457
                    TabIndex =8
                    Name ="AgeChampions"
                    ControlSource ="AgeChampions"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =755
                            Top =4395
                            Width =2070
                            Height =270
                            ForeColor =8388608
                            Name ="Text42"
                            Caption ="From results in age div only"
                            FontName ="Tahoma"
                            ControlTipText ="Determine the competitors points by totaling the points gained \015\012for event"
                                "s in the age division only.  Points from other divisions\015\012are not counted."
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    BorderWidth =3
                    Left =366
                    Top =5075
                    TabIndex =9
                    Name ="CompEvents"
                    ControlSource ="CompetitorEvents"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =619
                            Top =5019
                            Width =1785
                            Height =300
                            ForeColor =8388608
                            Name ="Text44"
                            Caption ="Competitor Events^"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =3276
                    Top =377
                    TabIndex =10
                    Name ="byEvent"
                    ControlSource ="EventResults"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3529
                            Top =321
                            Width =1785
                            Height =300
                            ForeColor =8388608
                            Name ="Text46"
                            Caption ="Event Results^"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =366
                    Top =5469
                    TabIndex =11
                    Name ="CompetitorList"
                    ControlSource ="CompetitorList"
                    DefaultValue ="No"
                    ControlTipText ="Useful list for manual entry of field event results."

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =619
                            Top =5416
                            Width =2040
                            Height =270
                            Name ="Text48"
                            Caption ="Competitor List (by Team)"
                            FontName ="Tahoma"
                            ControlTipText ="Useful list for manual entry of field event results."
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =1
                    OverlapFlags =215
                    Left =3276
                    Top =950
                    TabIndex =12
                    Name ="CurrentRecords"
                    ControlSource ="CurrentRecords"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3529
                            Top =897
                            Width =2025
                            Height =300
                            ForeColor =8388608
                            Name ="Current Records"
                            Caption ="Current Records^"
                            FontName ="Tahoma"
                            EventProcPrefix ="Current_Records"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =3276
                    Top =2129
                    TabIndex =13
                    Name ="CompResultsEvHouse"
                    ControlSource ="CompetitorResults"
                    StatusBarText ="Useful for determining relay teams."
                    DefaultValue ="No"
                    ControlTipText ="Useful for determining relay teams."

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3529
                            Top =2076
                            Width =1785
                            Height =510
                            Name ="Text52"
                            Caption ="Competitors Results (by Team / Event) *"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =288
                    Top =3384
                    Width =1029
                    Height =345
                    FontSize =8
                    FontWeight =400
                    TabIndex =14
                    Name ="All1"
                    Caption ="All"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =1536
                    Top =3384
                    Width =1029
                    Height =345
                    FontSize =8
                    FontWeight =400
                    TabIndex =15
                    Name ="None1"
                    Caption ="None"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =3348
                    Top =3390
                    Width =1029
                    Height =345
                    FontSize =8
                    FontWeight =400
                    TabIndex =16
                    Name ="All2"
                    Caption ="All"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =4540
                    Top =3390
                    Width =1029
                    Height =345
                    FontSize =8
                    FontWeight =400
                    TabIndex =17
                    Name ="None2"
                    Caption ="None"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =431
                    Top =6115
                    Width =1029
                    Height =345
                    FontSize =8
                    FontWeight =400
                    TabIndex =18
                    Name ="All3"
                    Caption ="All"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =1679
                    Top =6115
                    Width =1029
                    Height =345
                    FontSize =8
                    FontWeight =400
                    TabIndex =19
                    Name ="None3"
                    Caption ="None"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =3276
                    Top =2653
                    TabIndex =20
                    Name ="CompetitorPlaces"
                    ControlSource ="CompetitorPlace"
                    StatusBarText ="Orders competitor by final-level then place."
                    DefaultValue ="No"
                    ControlTipText ="Orders competitor by final-level then place."

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3527
                            Top =2598
                            Width =1470
                            Height =300
                            Name ="Text65"
                            Caption ="Competitor Places *"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =10371
                    Top =2310
                    Width =516
                    Height =225
                    TabIndex =21
                    Name ="CompetitorPlaces#"
                    ControlSource ="NumberOfRecords"
                    FontName ="Tahoma"
                    EventProcPrefix ="CompetitorPlaces_"
                    ControlTipText ="Reports that have an '*' will limit the rows returned."

                End
                Begin TextBox
                    OverlapFlags =215
                    Left =1725
                    Top =4110
                    Width =546
                    Height =225
                    TabIndex =22
                    Name ="Field70"
                    ControlSource ="AgeChampionNumber"
                    ControlTipText ="Enter the number of competitors to show in each age division."

                End
                Begin CheckBox
                    SpecialEffect =1
                    OverlapFlags =215
                    Left =3276
                    Top =1310
                    TabIndex =23
                    Name ="CurrentRecords-Mini"
                    ControlSource ="CurrentRecords-Mini"
                    DefaultValue ="No"
                    EventProcPrefix ="CurrentRecords_Mini"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3529
                            Top =1257
                            Width =2025
                            Height =300
                            Name ="Text72"
                            Caption ="Current Records (Mini)"
                            FontName ="Tahoma"
                            ControlTipText ="Report results affected by Selected Events"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =3276
                    Top =1671
                    TabIndex =24
                    Name ="RecordDate"
                    ControlSource ="Recordset"
                    StatusBarText ="Shows all records set on the specified date."
                    DefaultValue ="No"
                    ControlTipText ="Shows all records set on the specified date."

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3529
                            Top =1648
                            Width =1200
                            Height =270
                            Name ="Text74"
                            Caption ="Records set on "
                            FontName ="Tahoma"
                            LayoutCachedLeft =3529
                            LayoutCachedTop =1648
                            LayoutCachedWidth =4729
                            LayoutCachedHeight =1918
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9507
                    Top =780
                    Width =1239
                    Height =510
                    FontSize =7
                    FontWeight =400
                    TabIndex =25
                    Name ="PrintSelected"
                    Caption ="Print Selected Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9507
                    Top =1410
                    Width =1239
                    Height =510
                    FontSize =7
                    FontWeight =400
                    TabIndex =26
                    Name ="PrintOpen"
                    Caption ="Print All Open  Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    ListWidth =1105
                    Left =4750
                    Top =1617
                    Width =1255
                    Height =242
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="Date"
                    ControlSource ="RecordDate"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Format([Date],\"dd-mmm-yy\") AS Expr1, Records.Date FROM Records"
                        " ORDER BY Records.Date DESC;"
                    ColumnWidths ="855"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    Format ="Medium Date"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9606
                    Top =5421
                    Width =1119
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =28
                    HelpContextId =270
                    Name ="Button81"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =9225
                    Top =2316
                    Width =1050
                    Height =600
                    Name ="Text82"
                    Caption ="(*) Number of records to display"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =3276
                    Top =3009
                    TabIndex =29
                    Name ="Statistics-EventTimesOverallAsc"
                    ControlSource ="Fastest"
                    StatusBarText ="Orders competitors by result."
                    DefaultValue ="No"
                    EventProcPrefix ="Statistics_EventTim1"
                    ControlTipText ="Orders competitors by result."

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3527
                            Top =2958
                            Width =2265
                            Height =300
                            ForeColor =8388608
                            Name ="Q"
                            Caption ="Competitor Results*^"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =1
                    OverlapFlags =215
                    Left =427
                    Top =2050
                    TabIndex =30
                    Name ="Statistics-byPlace"
                    ControlSource ="Place"
                    DefaultValue ="No"
                    EventProcPrefix ="Statistics_byPlace"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =674
                            Top =1991
                            Width =2115
                            Height =300
                            Name ="Text88"
                            Caption ="Overall Res. by Place"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =9516
                    Top =3318
                    Width =1239
                    Height =720
                    FontSize =7
                    FontWeight =400
                    TabIndex =31
                    Name ="GenerateHTMLbut"
                    Caption ="Generate HTML for Selected Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9291
                    Top =4182
                    Width =1665
                    Height =510
                    ForeColor =8388608
                    Name ="Text90"
                    Caption ="(^) Has Web Page Facility"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =1
                    OverlapFlags =85
                    BorderWidth =3
                    Left =3276
                    Top =4802
                    TabIndex =32
                    Name ="Check92"
                    ControlSource ="=Yes"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =3540
                            Top =4758
                            Width =2325
                            Height =285
                            Name ="Label93"
                            Caption ="the \"Selected Events\""
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =3276
                    Top =5162
                    TabIndex =33
                    Name ="Check95"
                    ControlSource ="=Yes"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
                            Top =5118
                            Width =2325
                            Height =285
                            Name ="Label96"
                            Caption ="the \"Selected Teams\""
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BorderWidth =3
                    Left =3276
                    Top =5522
                    TabIndex =34
                    Name ="Check97"
                    ControlSource ="=Yes"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =87
                            Left =3540
                            Top =5478
                            Width =2535
                            Height =435
                            Name ="Label98"
                            Caption ="the \"Selected Events\" AND  \"Selected Teams\""
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =3276
                    Top =4038
                    Width =2955
                    Height =645
                    FontWeight =700
                    Name ="Label99"
                    Caption ="The type of checkbox against each report indicates that results are affected by:"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =427
                    Top =2470
                    TabIndex =35
                    Name ="Cumulative"
                    ControlSource ="Cumulative"
                    StatusBarText ="Useful data is available only if events have been numbered."
                    DefaultValue ="No"
                    ControlTipText ="Useful data is available only if events have been numbered."

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =674
                            Top =2411
                            Width =2115
                            Height =300
                            Name ="Label104"
                            Caption ="Cumulative Res. by Event #"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =346
                    Top =5828
                    TabIndex =36
                    Name ="Non-Participants"
                    ControlSource ="Non-Participants"
                    DefaultValue ="No"
                    EventProcPrefix ="Non_Participants"
                    ControlTipText ="Useful list for manual entry of field event results."

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =599
                            Top =5775
                            Width =1320
                            Height =270
                            Name ="Label106"
                            Caption ="Non-Participants"
                            FontName ="Tahoma"
                            ControlTipText ="Useful list for manual entry of field event results."
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =3
                    OverlapFlags =215
                    Left =495
                    Top =4748
                    TabIndex =37
                    Name ="AgeChampionAllDivs"
                    ControlSource ="AgeChampionAllDivs"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =735
                            Top =4695
                            Width =2145
                            Height =255
                            ForeColor =8388608
                            Name ="Label108"
                            Caption ="From results across all divs"
                            FontName ="Tahoma"
                            ControlTipText ="Determine the competitors points by totaling the points gained \015\012in ANY ag"
                                "e division."
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =285
                    Top =4110
                    Width =1380
                    Height =225
                    FontWeight =700
                    Name ="Label111"
                    Caption ="Age Champions"
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
Option Explicit

Private Sub AgeChamp_Click()
On Error GoTo Err_AgeChamp_Click

    Dim DocName As String

    DocName = "Statistics-AgeChampions"
    DoCmd.OpenReport DocName, A_PREVIEW

Exit_AgeChamp_Click:
    Exit Sub

Err_AgeChamp_Click:
    MsgBox Error$
    Resume Exit_AgeChamp_Click
    
End Sub

Private Sub All1_Click()

    Me![Overall].Value = True
    Me![byAge].Value = True
    Me![bySex].Value = True
    Me![bySexAge].Value = True
    Me![Statistics-byPlace].Value = True
    Me![Cumulative].Value = True

End Sub

Private Sub All2_Click()

    Me![byEvent].Value = True
    Me![CompResultsEvHouse].Value = True
    Me![CurrentRecords].Value = True
    Me![CurrentRecords-Mini].Value = True
    Me![RecordDate].Value = True
    Me![CompetitorPlaces].Value = True
    Me![Statistics-EventTimesOverallAsc] = True
    

End Sub

Private Sub All3_Click()

    Me![AgeChampions].Value = True
    Me![CompetitorList].Value = True
    Me![CompEvents].Value = True
    Me![Non-Participants].Value = True
    Me![AgeChampionAllDivs].Value = True
    
End Sub

Private Sub Button80_Click()
On Error GoTo Err_Button80_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Button80_Click:
    Exit Sub

Err_Button80_Click:
    MsgBox Error$
    Resume Exit_Button80_Click
    
End Sub

Private Sub ByAge_Click()
On Error GoTo Err_ByAge_Click

    Dim DocName As String

    DocName = "Statistics-Age"
    DoCmd.OpenReport DocName, A_PREVIEW

Exit_ByAge_Click:
    Exit Sub

Err_ByAge_Click:
    MsgBox Error$
    Resume Exit_ByAge_Click
    
End Sub

'Private Sub ByEvent_Click()
'On Error GoTo Err_ByEvent_Click

'    Dim DocName As String'

'    DocName = "Statistics-Event"
'    DoCmd.OpenReport DocName, A_PREVIEW

'Exit_ByEvent_Click:
'    Exit Sub

'Err_ByEvent_Click:
'    MsgBox Error$
'    Resume Exit_ByEvent_Click
    
'End Sub

'Private Sub BySex_Click()
'On Error GoTo Err_BySe_Click

'    Dim DocName As String

'    DocName = "Statistics-Sex"
'    DoCmd.OpenReport DocName, A_PREVIEW

'Exit_BySe_Click:
'    Exit Sub

'Err_BySe_Click:
'    MsgBox Error$
'    Resume Exit_BySe_Click
    

'End Sub

'Private Sub BySexAge_Click()
'On Error GoTo Err_BySexAge_Click

'    Dim DocName As String

'    DocName = "Statistics-Age/Sex"
'    DoCmd.OpenReport DocName, A_PREVIEW

'Exit_BySexAge_Click:
'    Exit Sub

'Err_BySexAge_Click:
'    MsgBox Error$
'    Resume Exit_BySexAge_Click
    
'End Sub

Private Sub Close_Click()
On Error GoTo Err_Close_Click


    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

'Private Sub CompEvents_Click()
'On Error GoTo Err_CompEvents_Click

'    Dim DocName As String

'    DocName = "Statistics-CompetitorEvents"
'    DoCmd.OpenReport DocName, A_PREVIEW

'Exit_CompEvents_Click:
'    Exit Sub

'Err_CompEvents_Click:
'    MsgBox Error$
'    Resume Exit_CompEvents_Click
    
'End Sub

'Private Sub CompEventsbyEventHou_Click()
'On Error GoTo Err_CompEventsbyEventHou_Click

'    Dim DocName As String

'    DocName = "Statistics-Competitor Results by House/Event"
'    DoCmd.OpenReport DocName, A_PREVIEW

'Exit_CompEventsbyEventHou_Click:
'    Exit Sub

'Err_CompEventsbyEventHou_Click:
'    MsgBox Error$
'    Resume Exit_CompEventsbyEventHou_Click
    

'End Sub

Private Sub Date_AfterUpdate()

    Me![RecordDate] = True
    DoCmd.RunCommand acCmdSaveRecord


End Sub

Private Sub Date_DblClick(Cancel As Integer)

    Me![RecordDate] = True
    Me![Date] = Date
    DoCmd.RunCommand acCmdSaveRecord

End Sub

'Private Sub EventResults_Click()
'On Error GoTo Err_EventResults_Click

'    Dim DocName As String

'    DocName = "Statistics-Event Places"
'    DoCmd.OpenReport DocName, A_PREVIEW

'Exit_EventResults_Click:
'    Exit Sub

'Err_EventResults_Click:
'    MsgBox Error$
'    Resume Exit_EventResults_Click
    

'End Sub

Private Sub Form_Activate()

    DoCmd.SelectObject A_FORM, "Statistical Reports", False
    DoCmd.Restore

End Sub

Private Sub Form_GotFocus()

    Stop
    DoCmd.Restore

End Sub

Private Sub GenerateHTMLbut_Click()
    Dim FileName As Variant, FileNum As Integer
    
    On Error GoTo GenerateHTMLbut_Click_Err

    Dim Response As Integer
    Close ' Close any opened files
    
    DoCmd.RunCommand acCmdSaveRecord
    
    ' *** Check that HTML Template Summary files have been set up ***
    FileName = DLookup("[TemplateFileSummary]", "MiscHTML")
    If IsNull(FileName) Then
        MsgBox ("You must set a HTML template summary file before you can generate the web page reports.  You can do this in HTML Utilities.")
        GoTo GenerateHTMLbut_Click_Exit
    Else
        FileNum = FreeFile
        On Error GoTo TemplateSumaryFile_Error
        Open FileName For Random As FileNum
        Close FileNum
        On Error GoTo GenerateHTMLbut_Click_Err
    End If

    ' *** Check that HTML Template files have been set up ***
    FileName = DLookup("[TemplateFile]", "MiscHTML")
    If IsNull(FileName) Then
        MsgBox ("You must set a HTML template file before you can generate the web page reports.  You can do this in HTML Utilities.")
        GoTo GenerateHTMLbut_Click_Exit
    Else
        FileNum = FreeFile
        On Error GoTo TemplateFile_Error
        Open FileName For Random As FileNum
        Close FileNum
        On Error GoTo GenerateHTMLbut_Click_Err
    End If
    
    ' *** Check output directory exists ***
    FileName = DLookup("[HTMLlocation]", "MiscHTML")
    If IsNull(FileName) Then
        MsgBox ("You must set a folder where the web pages will be saved before you can generate the web page reports.  You can do this in HTML Utilities.")
        GoTo GenerateHTMLbut_Click_Exit
    Else
        GlobalVariable = Dir(FileName & "\", vbDirectory)
        If GlobalVariable = "" Then
            Response = MsgBox("It appears that the folder where you want the web pages stored does not exist.  Do you want to create it now?", vbYesNo + vbInformation + vbDefaultButton1, "Create Folder Confirmation")
            If Response = 6 Then
                MkDir (FileName)
            Else
                MsgBox ("No web pages have been created.")
                GoTo GenerateHTMLbut_Click_Exit
            End If
        End If
        On Error GoTo GenerateHTMLbut_Click_Err
    End If
    
    Response = MsgBox("This action may overwrite web pags in the web directory " & DLookup("[HTMLlocation]", "MiscHTML") & ".  Do you want to continue?", vbInformation + vbYesNo, "Continue")
    If Response = 6 Then
        GlobalGenerateHTML = True
        
 '       Application.Echo False
        Call PrintPreviewReports("PREVIEW", True)
        
        DoCmd.RunMacro "ClosePleaseWait"
        MsgBox "Web pages have been generated.", vbInformation
    End If

GenerateHTMLbut_Click_Exit:
'    Application.Echo True
    Close ' Close all open files
    Exit Sub

GenerateHTMLbut_Click_Err:
    MsgBox (Error$)
    GoTo GenerateHTMLbut_Click_Exit

TemplateSumaryFile_Error:
    MsgBox ("The HTML template summary file you have selected is not valid.  Please check this file in HTML Utilities.  No web based reports will be generated until this is resolved.")
    GoTo GenerateHTMLbut_Click_Exit

TemplateFile_Error:
    MsgBox ("The HTML summary file you have selected is not valid.  Please check this file in HTML Utilities.")
    GoTo GenerateHTMLbut_Click_Exit

End Sub

Private Sub GoToLastPage(n As Variant)

    DoCmd.SelectObject A_REPORT, n, False
    
    'SendKeys "{f5}"
    'SendKeys "9999"
    'SendKeys "~"
    SendKeys "^{Right}"
    
End Sub

Private Sub None1_Click()

    Me![Overall].Value = False
    Me![byAge].Value = False
    Me![bySex].Value = False
    Me![bySexAge].Value = False
    Me![Statistics-byPlace].Value = False
    Me![Cumulative].Value = False
End Sub

Private Sub None2_Click()

    Me![byEvent].Value = False
    Me![CompResultsEvHouse].Value = False
    Me![CurrentRecords].Value = False
    Me![CurrentRecords-Mini].Value = False
    Me![RecordDate].Value = False
    Me![CompetitorPlaces].Value = False
    Me![Statistics-EventTimesOverallAsc] = False

End Sub

Private Sub None3_Click()

    Me![AgeChampions].Value = False
    Me![CompetitorList].Value = False
    Me![CompEvents].Value = False
    Me![Non-Participants].Value = False
    Me![AgeChampionAllDivs].Value = False
    
    
End Sub

Private Sub Overall_Click()
On Error GoTo Err_Overall_Click

    Dim DocName As String

    DocName = "Statistics-Overall"
    DoCmd.OpenReport DocName, A_PREVIEW

Exit_Overall_Click:
    Exit Sub

Err_Overall_Click:
    MsgBox Error$
    Resume Exit_Overall_Click
    
End Sub

Private Sub PrintOpen_Click()
On Error GoTo PrintOpen_Click_Err

  Dim Result As Variant

  DoCmd.RunCommand acCmdSaveRecord
  GlobalGenerateHTML = False
  Result = PrintOpenReports()

PrintOpen_Click_Exit:
  Exit Sub
  
PrintOpen_Click_Err:
  MsgBox Err.Description, vbCritical
  
End Sub

Private Sub PrintPreviewReports(Ty As String, Optional GenerateHTML)
On Error GoTo PrintPreviewReports_Err

    Dim Q As Variant, f As Form, T As Date
    
    If IsMissing(GenerateHTML) Then GenerateHTML = False
    
    DoCmd.RunCommand acCmdSaveRecord
    
    If GenerateHTML Then
      For Each f In Forms
        f.visible = False
      Next
    End If
    
    'DoCmd.RunCommand acCmdSaveRecord

    Dim DocName As String

    If Me![Cumulative].Value Then
      Call PrintPreviewCumulativeReport(Ty)
    End If
    
    If Me![Overall].Value Then
        DocName = "Statistics-Overall"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
          If GenerateHTML Then GoSub HandleHTML
          
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If
    
    If Me![byAge].Value Then
        DocName = "Statistics-Age"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
          If GenerateHTML Then GoSub HandleHTML
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If
    
    If Me![bySex].Value Then
        DocName = "Statistics-Sex"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
          If GenerateHTML Then GoSub HandleHTML
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    
    End If

    If Me![bySexAge].Value Then
        DocName = "Statistics-AgeGender"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
          If GenerateHTML Then GoSub HandleHTML
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If
    
    If Me![AgeChampions].Value Then
        
        DocName = "Statistics-AgeChampions"
        If Ty = "PREVIEW" Then
            If GenerateHTML Then
                Call ExportNamesHTML("agch")
            Else
                Call PreviewReport(DocName, acCmdPreviewTwoPages)
            End If
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If
    
    If Me.AgeChampionAllDivs Then
        
        DocName = "Statistics-AgeChampions-AcrossAllDivisions"
        If Ty = "PREVIEW" Then
            If GenerateHTML Then
                Call ExportNamesHTML("agca")
            Else
                Call PreviewReport(DocName, acCmdPreviewTwoPages)
            End If
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![byEvent].Value Then
        DocName = "Statistics-Event"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
          If GenerateHTML Then GoSub HandleHTML
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![CompEvents].Value Then
        DocName = "Statistics-CompetitorEvents"
        If Ty = "PREVIEW" Then
            If GenerateHTML Then
                Call ExportNamesHTML("coev")
            Else
                Call PreviewReport(DocName, acCmdPreviewTwoPages)
            End If
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![CompetitorList].Value Then
        DocName = "CompetitorList-ByTeamAge"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![CurrentRecords].Value Then
        DocName = "RecordDetails-Current"
        If Ty = "PREVIEW" Then
            If GenerateHTML Then
                Call ExportNamesHTML("rh")
            Else
                Call PreviewReport(DocName, acCmdPreviewTwoPages)
            End If
        Else
          DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![CurrentRecords-Mini].Value Then
        DocName = "RecordDetailsMini-Current"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If
    
    If Me![RecordDate].Value Then
        DocName = "RecordDetails-ByDay"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![CompetitorPlaces].Value Then
        DocName = "Statistics-EventPlaces"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![CompResultsEvHouse].Value Then
        DocName = "Statistics-CompetitorResultsbyEventTeam"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![Statistics-EventTimesOverallAsc].Value Then
        DocName = "Statistics-EventTimesOverallAsc"
        If Ty = "PREVIEW" Then
            If GenerateHTML Then
                Call ExportNamesHTML("etoa")
            Else
                Call PreviewReport(DocName, acCmdPreviewTwoPages)
            End If
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![Statistics-byPlace].Value Then
        DocName = "Statistics-byPlace"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If

    If Me![Non-Participants].Value Then
      DocName = "Misc-Non Participators"
        If Ty = "PREVIEW" Then
          Call PreviewReport(DocName, acCmdPreviewTwoPages)
        Else
            DoCmd.OpenReport DocName, A_NORMAL
        End If
    End If
      
PrintPreviewReports_Exit:
  
    If GenerateHTML Then
      For Each f In Forms
        f.visible = True
      Next
    End If
  
  Exit Sub
  
PrintPreviewReports_Err:
    If Err.Number = 2501 Then ' Any error other than No Report Data error
      Resume Next
    Else
      MsgBox Error$
      GoTo PrintPreviewReports_Exit
    End If

HandleHTML:
    
    PleaseWaitMsg = "Finalising HTML for """ & DocName & """.  Please wait..."
    DoCmd.RunMacro "ShowPleaseWait"

    DoCmd.SelectObject acReport, DocName, False
    DoCmd.Maximize
    ' As reports use .Format functions to generate code you need to move to last page of report to generate all entries
    ' In newer versions of Access can just send Ctrl +  right arrow or End
    'SendKeys "{End}", True
    SendKeys "^{Right}"

    T = Timer
    SysCmd acSysCmdSetStatus, "Waiting for report process to finalise..."
    ' Wait 1 second
    Do While Timer < T + 1
      DoEvents
    Loop
    DoCmd.Close acReport, DocName
    SysCmd acSysCmdClearStatus
    
  Return
  
End Sub

Private Sub PrintSelected_Click()
On Error GoTo PrintSelected_Click_Err

  DoCmd.RunCommand acCmdSaveRecord
  GlobalGenerateHTML = False
  PrintPreviewReports ("PRINT")

PrintSelected_Click_Exit:
  Exit Sub
  
PrintSelected_Click_Err:
  MsgBox Err.Description, vbCritical
  
End Sub

Private Sub Selected_Click()
On Error GoTo Err_Selected_Click
  
  DoCmd.RunCommand acCmdSaveRecord
  
  PleaseWaitMsg = "Opening selected reports ..."
  DoCmd.RunMacro "ShowpleaseWait"
  
  GlobalGenerateHTML = False
  Call PrintPreviewReports("PREVIEW")
  DoCmd.RunMacro "ClosePleaseWait"

  DoCmd.OpenForm "ReportsPopUp"
  
Exit_Selected_Click:
    Exit Sub

Err_Selected_Click:
    MsgBox Error$
    Resume Exit_Selected_Click
    
End Sub

Private Sub PrintPreviewCumulativeReport(Ty As String)
On Error GoTo PrintPreviewCumulativeReport_Err

  Dim DocName As String, rs As Recordset, brs As Recordset
  Dim OldTeam As Variant, OldEvent As Variant, Points As Integer
  Dim MinEventNum As Variant, MaxEventNum As Variant, EventNum As Variant

  DocName = "Statistics-PointByEventNumber"
  
  Set brs = CurrentDb.OpenRecordset("Points by EventNumber and Team", dbOpenDynaset)
  Set rs = CurrentDb.OpenRecordset("Points-House-Cumulative", dbOpenDynaset)
  
  MinEventNum = DMin("[E_Num]", "Points by EventNumber and Team")
  MaxEventNum = DMax("[E_Num]", "Points by EventNumber and Team")
  
  If IsNull(MinEventNum) Or IsNull(MaxEventNum) Or MinEventNum = MaxEventNum Then
    MsgBox "No event numbers have been allocated.  Please allocate event numbers before running the cumulative report.", vbInformation
    
  ElseIf Not brs.BOF Then
    DoCmd.SetWarnings False
    'DoCmd.RunSQL "Delete * from [Points-House-Cumulative]"
    DoCmd.RunSQL "UPDATE [Points-House-Cumulative] SET [Points-House-Cumulative].Flag = False"

    DoCmd.SetWarnings True
    If rs.BOF Then ' Ensure there is at least on record in table
      'rs.AddNew
      
      'rs!Flag = False
      'rs.Update
    End If
    'rs.MoveFirst
    
    OldEvent = MinEventNum
    
    While Not brs.EOF
    
      OldTeam = brs!H_ID
      
      For EventNum = MinEventNum To MaxEventNum
        If brs.EOF Then GoTo ExitLoop
        While (brs!E_Num = EventNum) And (brs!H_ID = OldTeam)
          
          Points = Points + ConvertNullToZero(brs!Points)
          brs.MoveNext
          If brs.EOF Then GoTo ExitLoop
        Wend
ExitLoop:
        
        If rs.EOF Then
          rs.AddNew
        Else
          rs.Edit
          
        End If
        rs!H_ID = OldTeam
        rs!E_Number = EventNum
        rs!Points = Points
        rs!Flag = True
        rs.Update
        If Not rs.EOF Then rs.MoveNext
      Next
      Points = 0

    Wend
    
    If Ty = "PREVIEW" Then
        DoCmd.OpenReport DocName, A_PREVIEW
        DoCmd.Maximize
        DoCmd.RunCommand acCmdPreviewTwoPages
    Else
        DoCmd.OpenReport DocName, A_NORMAL
    End If
  Else
    MsgBox "No cumulative results to show.", vbInformation
  End If

PrintPreviewCumulativeReport_Exit:
  Exit Sub
  
PrintPreviewCumulativeReport_Err:
  MsgBox "An error has occurred in [PrintPreviewCumulativeReport]: " & Err.Description, vbCritical
  Resume PrintPreviewCumulativeReport_Exit
  
End Sub
