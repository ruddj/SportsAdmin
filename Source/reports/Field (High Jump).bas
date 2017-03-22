Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10716
    ItemSuffix =141
    Left =750
    Top =90
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xb3922e86eee5e140
    End
    RecordSource ="Field Events"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x3702000037020000370200003702000000000000dc2900004002000001000000 ,
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
            ControlSource ="FullName"
        End
        Begin PageHeader
            Height =1945
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10716
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
                    Width =10656
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
                    Left =8054
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
                            Left =7370
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
                    Left =9529
                    Top =680
                    Width =1117
                    Height =287
                    FontSize =11
                    TabIndex =3
                    Name ="Field82"
                    ControlSource ="=IIf([Sex]=\"F\",\"Girls\",IIf([Sex]=\"M\",\"Boys\",\"Mixed\"))"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            TextFontFamily =18
                            Left =8844
                            Top =682
                            Width =615
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
                    Width =4704
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
                    Width =10715
                    Name ="Line75"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =170
                    Top =1611
                    Width =3000
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
                    Left =3409
                    Top =1588
                    Width =1020
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
                    Left =9927
                    Top =1588
                    Width =720
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text89"
                    Caption ="PLACE"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =8957
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
                    Top =1077
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
                            Top =1077
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
                    Left =10085
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
                    Width =4697
                    Height =331
                    FontSize =11
                    TabIndex =9
                    Name ="Field93"
                    ControlSource ="RecHolder"
                    FontName ="Times New Roman"

                End
                Begin Line
                    LineSlant = NotDefault
                    Top =1928
                    Width =10647
                    Name ="Line77"
                End
                Begin Line
                    Left =4762
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line120"
                End
                Begin Line
                    Left =5045
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line121"
                End
                Begin Line
                    Left =5329
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line122"
                End
                Begin Line
                    Left =5613
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line123"
                End
                Begin Line
                    Left =5897
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line124"
                End
                Begin Line
                    Left =6181
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line125"
                End
                Begin Line
                    Left =6465
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line126"
                End
                Begin Line
                    Left =6749
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line127"
                End
                Begin Line
                    Left =7033
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line128"
                End
                Begin Line
                    Left =7317
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line129"
                End
                Begin Line
                    Left =7601
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line130"
                End
                Begin Line
                    Left =7885
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line131"
                End
                Begin Line
                    Left =8169
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line132"
                End
                Begin Line
                    Left =8453
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line133"
                End
                Begin Line
                    Left =8737
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line134"
                End
                Begin Line
                    Left =9021
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line135"
                End
                Begin Line
                    Left =9305
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line136"
                End
                Begin Line
                    Left =9589
                    Top =1474
                    Width =0
                    Height =471
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Left =4423
                    Top =1484
                    Width =0
                    Height =441
                    Name ="Line139"
                End
                Begin Line
                    BorderWidth =2
                    Left =9864
                    Top =1484
                    Width =0
                    Height =441
                    Name ="Line140"
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
            Height =591
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =657
                    Top =144
                    Width =2543
                    Height =340
                    FontSize =10
                    Name ="Field21"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3285
                    Top =144
                    Width =1082
                    Height =332
                    FontSize =10
                    TabIndex =1
                    Name ="Field23"
                    ControlSource ="H_Code"

                End
                Begin Line
                    Left =5
                    Width =0
                    Height =576
                    Name ="Line56"
                End
                Begin Line
                    Left =510
                    Width =0
                    Height =576
                    Name ="Line57"
                End
                Begin Line
                    Left =3231
                    Width =0
                    Height =576
                    Name ="Line58"
                End
                Begin Line
                    BorderWidth =2
                    Left =4423
                    Width =0
                    Height =576
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =2
                    Left =9864
                    Width =0
                    Height =576
                    Name ="Line61"
                End
                Begin Line
                    Left =5
                    Top =576
                    Width =10647
                    Name ="Line62"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =9977
                    Top =170
                    Width =617
                    Height =332
                    FontSize =10
                    TabIndex =2
                    Name ="Place"
                    ControlSource ="F_Place"

                End
                Begin Line
                    Left =10658
                    Width =0
                    Height =576
                    Name ="Line81"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =56
                    Top =113
                    Width =329
                    Height =340
                    FontSize =10
                    TabIndex =3
                    Name ="Field12"
                    ControlSource ="=1"

                End
                Begin Line
                    Left =4762
                    Width =0
                    Height =576
                    Name ="Line82"
                End
                Begin Line
                    Left =5045
                    Width =0
                    Height =576
                    Name ="Line83"
                End
                Begin Line
                    Left =5329
                    Width =0
                    Height =576
                    Name ="Line84"
                End
                Begin Line
                    Left =5613
                    Width =0
                    Height =576
                    Name ="Line85"
                End
                Begin Line
                    Left =5897
                    Width =0
                    Height =576
                    Name ="Line86"
                End
                Begin Line
                    Left =6181
                    Width =0
                    Height =576
                    Name ="Line88"
                End
                Begin Line
                    Left =6465
                    Width =0
                    Height =576
                    Name ="Line89"
                End
                Begin Line
                    Left =6749
                    Width =0
                    Height =576
                    Name ="Line90"
                End
                Begin Line
                    Left =7033
                    Width =0
                    Height =576
                    Name ="Line91"
                End
                Begin Line
                    Left =7317
                    Width =0
                    Height =576
                    Name ="Line92"
                End
                Begin Line
                    Left =7601
                    Width =0
                    Height =576
                    Name ="Line93"
                End
                Begin Line
                    Left =7885
                    Width =0
                    Height =576
                    Name ="Line94"
                End
                Begin Line
                    Left =8169
                    Width =0
                    Height =576
                    Name ="Line95"
                End
                Begin Line
                    Left =8453
                    Width =0
                    Height =576
                    Name ="Line96"
                End
                Begin Line
                    Left =8737
                    Width =0
                    Height =576
                    Name ="Line97"
                End
                Begin Line
                    Left =9021
                    Width =0
                    Height =576
                    Name ="Line98"
                End
                Begin Line
                    Left =9305
                    Width =0
                    Height =576
                    Name ="Line99"
                End
                Begin Line
                    Left =9589
                    Width =0
                    Height =576
                    Name ="Line100"
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
                    Left =9581
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Width =9591
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
                    Width =10716
                    Name ="Line87"
                End
            End
        End
    End
End
