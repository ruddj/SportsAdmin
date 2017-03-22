Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =14631
    ItemSuffix =102
    Left =285
    Top =165
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x6c3b59d2cfe5e140
    End
    RecordSource ="Lanes Limited Report"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x3702000037020000370200003702000000000000273900001903000001000000 ,
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
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="Sex"
        End
        Begin BreakLevel
            ControlSource ="FLevSub"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Heat"
        End
        Begin BreakLevel
            ControlSource ="LaneSub"
        End
        Begin PageHeader
            Height =1803
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
                    Left =8224
                    Top =680
                    Width =801
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
                            Width =1095
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
                    Left =9699
                    Top =680
                    Width =952
                    Height =287
                    FontSize =11
                    TabIndex =3
                    Name ="Field82"
                    ControlSource ="=IIf([Sex]=\"F\",\"Girls\",\"Boys\")"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =9014
                            Top =682
                            Width =1035
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
                    TextAlign =3
                    TextFontFamily =18
                    Left =9077
                    Top =1020
                    Width =1022
                    Height =331
                    FontSize =11
                    TabIndex =5
                    Name ="Field91"
                    ControlSource ="record"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =1152
                    Top =1076
                    Width =1689
                    Height =332
                    FontSize =11
                    TabIndex =6
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
                    TabIndex =7
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
                Begin TextBox
                    TextFontFamily =18
                    Left =10205
                    Top =1020
                    Width =512
                    Height =331
                    FontSize =11
                    TabIndex =8
                    Name ="Field89"
                    ControlSource ="units"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =4191
                    Top =1020
                    Width =4832
                    Height =331
                    FontSize =11
                    TabIndex =9
                    Name ="Field93"
                    ControlSource ="RecHolder"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextFontFamily =18
                    Left =11459
                    Top =680
                    Width =2887
                    Height =287
                    FontSize =11
                    TabIndex =10
                    Name ="Text100"
                    ControlSource ="=IIf([E_Time]>=1,Format([E_Time],\"d-mmm h:nn am/pm\"),Format([E_Time],\"h:nn am"
                        "/pm\"))"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =10714
                            Top =682
                            Width =1095
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Label101"
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
            BreakLevel =4
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =793
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =173
                    Top =56
                    Width =794
                    Height =340
                    FontSize =10
                    TabIndex =1
                    Name ="Field12"
                    ControlSource ="LaneSub"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =1325
                    Top =56
                    Width =3458
                    Height =340
                    FontSize =10
                    TabIndex =2
                    Name ="Field21"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =5069
                    Top =56
                    Width =1877
                    Height =332
                    FontSize =10
                    TabIndex =3
                    Name ="Field23"
                    ControlSource ="H_Code"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7169
                    Top =82
                    Width =932
                    Height =332
                    FontSize =10
                    Name ="Place"
                    ControlSource ="F_Place"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =8188
                    Top =82
                    Width =1517
                    Height =332
                    FontSize =10
                    TabIndex =4
                    Name ="Result"
                    ControlSource ="cResult"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =34
                    Left =2665
                    Top =453
                    Width =7103
                    Height =230
                    FontSize =10
                    TabIndex =5
                    Name ="Field69"
                    ControlSource ="Memo"

                End
                Begin Rectangle
                    BackStyle =0
                    Width =1195
                    Height =445
                    Name ="Box70"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =1209
                    Width =3790
                    Height =445
                    Name ="Box71"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =4998
                    Width =2140
                    Height =445
                    Name ="Box72"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =7147
                    Width =1015
                    Height =445
                    Name ="Box73"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =8167
                    Width =1600
                    Height =445
                    Name ="Box74"
                End
                Begin Label
                    TextFontFamily =34
                    Left =1247
                    Top =453
                    Width =1245
                    Height =225
                    FontSize =10
                    Name ="Text75"
                    Caption ="Competitors:"
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
                    Left =13606
                    Width =1011
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Width =13386
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
                    Width =14631
                    Name ="Line87"
                End
            End
        End
    End
End
