Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10714
    ItemSuffix =94
    Left =2835
    Top =15
    OnNoData ="[Event Procedure]"
    Filter ="([R_Code] = 3 AND ([Status]=0 OR [Status]=1 OR [Status]=2 OR [Status]=3))"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x7f52c6ede0f3e140
    End
    RecordSource ="Field Events"
    Caption ="Unlimited Lanes List"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x3702000037020000370200009602000000000000da2900005f01000001000000 ,
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="HE_Code"
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
            ControlSource ="Heat"
        End
        Begin BreakLevel
            ControlSource ="H_Code"
        End
        Begin BreakLevel
            ControlSource ="FullName"
        End
        Begin PageHeader
            Height =680
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10656
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
                    Width =10596
                    Name ="Line74"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1148
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =1020
                    Width =669
                    Height =272
                    FontSize =11
                    Name ="Field65"
                    ControlSource ="E_Number"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Left =56
                            Top =4
                            Width =1005
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Text66"
                            Caption ="Event #:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =8280
                    Width =801
                    Height =287
                    FontSize =11
                    TabIndex =1
                    Name ="Field49"
                    ControlSource ="Age"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =7596
                            Top =2
                            Width =1095
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Text50"
                            Caption ="AGE:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =9755
                    Width =952
                    Height =287
                    FontSize =11
                    TabIndex =2
                    Name ="Field51"
                    ControlSource ="=IIf([Sex]=\"F\",\"Girls\",IIf([Sex]=\"M\",\"Boys\",\"Mixed\"))"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =9070
                            Top =2
                            Width =1035
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Text52"
                            Caption ="SEX:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =2719
                    Width =4869
                    Height =302
                    FontSize =11
                    TabIndex =3
                    BorderColor =16777215
                    Name ="Field45"
                    ControlSource ="ET_Des"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Left =1757
                            Top =6
                            Width =870
                            Height =315
                            FontSize =11
                            FontWeight =700
                            BorderColor =16777215
                            Name ="Text46"
                            Caption ="EVENT:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Line
                    Left =56
                    Top =340
                    Width =10595
                    Name ="Line75"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =192
                    Top =817
                    Width =846
                    Height =272
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text38"
                    Caption ="LANE"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =1140
                    Top =817
                    Width =3840
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text39"
                    Caption ="COMPETITOR"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5215
                    Top =796
                    Width =1875
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text40"
                    Caption ="TEAM"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =7200
                    Top =796
                    Width =1005
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text41"
                    Caption ="PLACE"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =8451
                    Top =793
                    Width =870
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text42"
                    Caption ="RESULT"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =9133
                    Top =340
                    Width =1022
                    Height =331
                    FontSize =11
                    TabIndex =4
                    Name ="Field53"
                    ControlSource ="record"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =1208
                    Top =396
                    Width =1689
                    Height =332
                    FontSize =11
                    TabIndex =5
                    Name ="Field63"
                    ControlSource ="FLevSub"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Left =56
                            Top =396
                            Width =1140
                            Height =330
                            FontSize =11
                            Name ="Text64"
                            Caption ="Final Level:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =3571
                    Top =396
                    Width =354
                    Height =332
                    FontSize =11
                    TabIndex =6
                    Name ="Field47"
                    ControlSource ="Heat"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Left =3004
                            Top =396
                            Width =570
                            Height =375
                            FontSize =11
                            Name ="Text48"
                            Caption ="Heat:"
                            FontName ="Times New Roman"
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =9694
                    Top =796
                    Width =870
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text80"
                    Caption ="POINTS"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =56
                    Top =1133
                    Width =10647
                    Name ="Line83"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =10261
                    Top =340
                    Width =437
                    Height =331
                    FontSize =11
                    TabIndex =7
                    Name ="Field89"
                    ControlSource ="units"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =4247
                    Top =340
                    Width =4832
                    Height =331
                    FontSize =11
                    TabIndex =8
                    Name ="Field93"
                    ControlSource ="RecHolder"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =355
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =1262
                    Width =3653
                    Height =265
                    FontSize =10
                    Name ="Field21"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =5101
                    Width =2117
                    Height =265
                    FontSize =10
                    TabIndex =1
                    Name ="Field23"
                    ControlSource ="H_Code"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4988
                    Width =0
                    Height =351
                    Name ="Line58"
                End
                Begin Line
                    Left =8190
                    Width =0
                    Height =351
                    Name ="Line60"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7497
                    Width =617
                    Height =265
                    FontSize =10
                    TabIndex =2
                    Name ="Place"
                    ControlSource ="F_Place"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =8277
                    Width =1262
                    Height =265
                    FontSize =10
                    TabIndex =3
                    Name ="Result"
                    ControlSource ="cResult"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1133
                    Width =0
                    Height =351
                    Name ="Line76"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =56
                    Width =0
                    Height =351
                    Name ="Line77"
                End
                Begin Line
                    Left =9637
                    Width =0
                    Height =351
                    Name ="Line78"
                End
                Begin Line
                    Left =10714
                    Width =0
                    Height =351
                    Name ="Line79"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =9706
                    Width =857
                    Height =265
                    FontSize =10
                    TabIndex =4
                    Name ="Points"
                    ControlSource ="cPoints"

                End
                Begin Line
                    Left =7314
                    Width =0
                    Height =351
                    Name ="Line82"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =56
                    Top =340
                    Width =10647
                    Name ="Line86"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =226
                    Width =694
                    Height =280
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="Field66"
                    ControlSource ="=1"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =170
            BreakLevel =1
            Name ="GroupFooter2"
        End
        Begin PageFooter
            Height =503
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Top =113
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =113
                    Width =9426
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
                    Top =113
                    Width =10596
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

  Call UnLimitedLanes_NoData
  Cancel = True
  
End Sub
