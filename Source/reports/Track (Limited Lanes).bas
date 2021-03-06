﻿Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =14631
    ItemSuffix =84
    Left =1410
    Top =195
    OnNoData ="[Event Procedure]"
    Filter ="([R_Code] = 1 AND ([Status]=0 OR [Status]=1 OR [Status]=2 OR [Status]=3))"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x7f2ce18feee5e140
    End
    RecordSource ="Lanes Limited Report"
    Caption ="Limited Lanes Detailed List"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x4a02000037020000e5010000370200000000000027390000d701000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
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
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="E_Number"
        End
        Begin BreakLevel
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="Sex"
        End
        Begin BreakLevel
            ControlSource ="F_Lev"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Heat"
        End
        Begin BreakLevel
            ControlSource ="LaneSub"
        End
        Begin PageHeader
            Height =1829
            Name ="PageHeader0"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =10770
                    Top =1360
                    Width =3795
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Text69"
                    Caption ="Timing Tape / Notes"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =18
                    Width =14631
                    Height =450
                    FontSize =16
                    FontWeight =700
                    Name ="CarnivalTitle"
                    ControlSource ="=DLookUp(\"[CarnivalTitle]\",\"Miscellaneous\")"
                    FontName ="times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =3
                    BorderLineStyle =3
                    Left =56
                    Top =566
                    Width =14571
                    Name ="Line74"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =964
                    Top =680
                    Width =669
                    Height =272
                    FontSize =11
                    TabIndex =1
                    Name ="Field78"
                    ControlSource ="E_Number"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Top =684
                            Width =1005
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Text79"
                            Caption ="Event #:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =8164
                    Top =680
                    Width =861
                    Height =287
                    FontSize =11
                    TabIndex =2
                    Name ="Field80"
                    ControlSource ="Age"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =7540
                            Top =682
                            Width =615
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Text81"
                            Caption ="AGE:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =9639
                    Top =680
                    Width =1012
                    Height =287
                    FontSize =11
                    TabIndex =3
                    Name ="Field82"
                    ControlSource ="=IIf([Sex]=\"F\",\"Girls\",IIf([Sex]=\"M\",\"Boys\",\"Mixed\"))"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =9014
                            Top =682
                            Width =570
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Text83"
                            Caption ="SEX:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =2663
                    Top =680
                    Width =4869
                    Height =302
                    FontSize =11
                    TabIndex =4
                    BorderColor =16777215
                    Name ="Field84"
                    ControlSource ="ET_Des"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Left =1701
                            Top =686
                            Width =870
                            Height =315
                            FontSize =11
                            FontWeight =700
                            BorderColor =16777215
                            Name ="Text85"
                            Caption ="EVENT:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Line
                    Top =1020
                    Width =14630
                    Name ="Line75"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =136
                    Top =1497
                    Width =846
                    Height =272
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text86"
                    Caption ="LANE"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =1084
                    Top =1497
                    Width =3840
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text87"
                    Caption ="COMPETITOR"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5159
                    Top =1476
                    Width =1875
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text88"
                    Caption ="TEAM"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =7144
                    Top =1476
                    Width =1005
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text89"
                    Caption ="PLACE"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =8395
                    Top =1473
                    Width =870
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text90"
                    Caption ="RESULT"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =1152
                    Top =1076
                    Width =1689
                    Height =332
                    FontSize =11
                    TabIndex =5
                    Name ="Field92"
                    ControlSource ="FLevSub"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Top =1076
                            Width =1140
                            Height =330
                            FontSize =11
                            Name ="Text93"
                            Caption ="Final Level:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =3515
                    Top =1076
                    Width =354
                    Height =332
                    FontSize =11
                    TabIndex =6
                    Name ="Field94"
                    ControlSource ="Heat"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Left =2948
                            Top =1076
                            Width =570
                            Height =375
                            FontSize =11
                            Name ="Text95"
                            Caption ="Heat:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Line
                    Top =1814
                    Width =9792
                    Name ="Line77"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =4005
                    Top =1110
                    Width =5721
                    Height =270
                    FontSize =9
                    TabIndex =7
                    Name ="Field79"
                    ControlSource ="=[RecHolder] & \" \" & [record] & \" \" & [units]"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =11459
                    Top =680
                    Width =2542
                    Height =287
                    FontSize =11
                    TabIndex =8
                    Name ="Text82"
                    ControlSource ="=IIf([E_Time]>=1,Format([E_Time],\"d-mmm h:nn am/pm\"),Format([E_Time],\"h:nn am"
                        "/pm\"))"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =10714
                            Top =682
                            Width =750
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Label83"
                            Caption ="TIME:"
                            FontName ="Times New Roman"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            BreakLevel =5
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =471
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =170
                    Top =56
                    Width =794
                    Height =340
                    FontSize =10
                    Name ="Field12"
                    ControlSource ="LaneSub"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =1322
                    Top =56
                    Width =3458
                    Height =340
                    FontSize =10
                    TabIndex =1
                    Name ="Field21"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =5066
                    Top =56
                    Width =1877
                    Height =332
                    FontSize =10
                    TabIndex =2
                    Name ="Field23"
                    ControlSource ="H_Code"

                End
                Begin Line
                    Left =5
                    Width =0
                    Height =471
                    Name ="Line56"
                End
                Begin Line
                    Left =1157
                    Width =0
                    Height =471
                    Name ="Line57"
                End
                Begin Line
                    Left =4901
                    Width =0
                    Height =471
                    Name ="Line58"
                End
                Begin Line
                    Left =7061
                    Width =0
                    Height =471
                    Name ="Line59"
                End
                Begin Line
                    Left =8069
                    Width =0
                    Height =471
                    Name ="Line60"
                End
                Begin Line
                    Left =9797
                    Width =0
                    Height =471
                    Name ="Line61"
                End
                Begin Line
                    Top =453
                    Width =9792
                    Name ="Line62"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7166
                    Top =82
                    Width =857
                    Height =332
                    FontSize =10
                    TabIndex =3
                    Name ="Place"
                    ControlSource ="F_Place"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =8185
                    Top =82
                    Width =1592
                    Height =332
                    FontSize =10
                    TabIndex =4
                    Name ="Result"
                    ControlSource ="cResult"

                End
            End
        End
        Begin PageFooter
            Height =390
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =13546
                    Width =1071
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Width =13551
                    Height =390
                    FontSize =11
                    FontWeight =700
                    Name ="Field88"
                    ControlSource ="=DLookUp(\"[CarnivalFooter]\",\"Miscellaneous\")"
                    FontName ="Times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =2
                    BorderLineStyle =3
                    Width =14616
                    Name ="Line87"
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
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)

  Call LimitedLanes_NoData
  Cancel = True

End Sub
