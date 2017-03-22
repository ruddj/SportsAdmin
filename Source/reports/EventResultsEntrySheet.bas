Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10771
    ItemSuffix =68
    Left =795
    Top =330
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x365f747beee5e140
    End
    RecordSource ="EventEntrryLists"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200008d01000037020000d0020000000000007d140000bd09000000000000 ,
        0x020000003702000000000000a20700000100000000000000
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
            ControlSource ="F_Lev_Sub"
        End
        Begin BreakLevel
            ControlSource ="Heat"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =7710
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =235
                    Top =560
                    Width =531
                    Height =242
                    FontWeight =700
                    Name ="Field49"
                    ControlSource ="Age"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =232
                            Top =340
                            Width =570
                            Height =240
                            Name ="Text50"
                            Caption ="Age"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =856
                    Top =564
                    Width =1327
                    Height =242
                    FontWeight =700
                    TabIndex =1
                    Name ="Field51"
                    ControlSource ="=IIf([Sex]=\"F\",\"Girls\",IIf([Sex]=\"M\",\"Boys\",\"Mixed\"))"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =858
                            Top =340
                            Width =1320
                            Height =225
                            Name ="Text52"
                            Caption ="Sex"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3860
                    Top =564
                    Width =847
                    Height =242
                    FontWeight =700
                    TabIndex =2
                    Name ="Field12"
                    ControlSource ="Heat"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =3860
                            Top =340
                            Width =870
                            Height =210
                            Name ="Text13"
                            Caption ="Division"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =2273
                    Top =564
                    Width =1327
                    Height =242
                    FontWeight =700
                    TabIndex =4
                    Name ="Field60"
                    ControlSource ="F_Lev_Sub"

                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =34
                            Left =2275
                            Top =340
                            Width =1320
                            Height =225
                            Name ="Text61"
                            Caption ="Final Level"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    Left =113
                    Top =1129
                    Width =4649
                    Height =450
                    Name ="Box64"
                End
                Begin Line
                    Left =1644
                    Top =1129
                    Width =0
                    Height =450
                    Name ="Line65"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =281
                    Top =1189
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text102"
                    Caption ="0"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    OldBorderStyle =0
                    BorderWidth =2
                    Left =113
                    Top =850
                    Width =4649
                    Height =273
                    BackColor =0
                    Name ="Box37"
                End
                Begin Line
                    Left =170
                    Top =558
                    Width =4649
                    Name ="Line14"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =2044
                    Width =4649
                    Height =450
                    Name ="Box15"
                End
                Begin Line
                    Left =1649
                    Top =2044
                    Width =0
                    Height =450
                    Name ="Line16"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =2104
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text17"
                    Caption ="2"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =2498
                    Width =4649
                    Height =450
                    Name ="Box18"
                End
                Begin Line
                    Left =1649
                    Top =2498
                    Width =0
                    Height =450
                    Name ="Line19"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =2558
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text20"
                    Caption ="3"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =2944
                    Width =4649
                    Height =450
                    Name ="Box21"
                End
                Begin Line
                    Left =1649
                    Top =2944
                    Width =0
                    Height =450
                    Name ="Line22"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =3004
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text23"
                    Caption ="4"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =3398
                    Width =4649
                    Height =465
                    Name ="Box24"
                End
                Begin Line
                    Left =1649
                    Top =3398
                    Width =0
                    Height =450
                    Name ="Line25"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =3458
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text26"
                    Caption ="5"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =3858
                    Width =4649
                    Height =450
                    Name ="Box27"
                End
                Begin Line
                    Left =1649
                    Top =3858
                    Width =0
                    Height =450
                    Name ="Line28"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =3918
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text29"
                    Caption ="6"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =4312
                    Width =4649
                    Height =450
                    Name ="Box30"
                End
                Begin Line
                    Left =1649
                    Top =4312
                    Width =0
                    Height =450
                    Name ="Line31"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =4372
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text32"
                    Caption ="7"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =4758
                    Width =4649
                    Height =450
                    Name ="Box33"
                End
                Begin Line
                    Left =1649
                    Top =4758
                    Width =0
                    Height =450
                    Name ="Line34"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =4818
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text35"
                    Caption ="8"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =5212
                    Width =4649
                    Height =450
                    Name ="Box36"
                End
                Begin Line
                    Left =1649
                    Top =5212
                    Width =0
                    Height =450
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =5272
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text38"
                    Caption ="9"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =5658
                    Width =4649
                    Height =465
                    Name ="Box39"
                End
                Begin Line
                    Left =1649
                    Top =5658
                    Width =0
                    Height =450
                    Name ="Line40"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =5718
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text43"
                    Caption ="10"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =6128
                    Width =4649
                    Height =420
                    Name ="Box48"
                End
                Begin Line
                    Left =1649
                    Top =6098
                    Width =0
                    Height =450
                    Name ="Line49"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =6158
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text51"
                    Caption ="11"
                    FontName ="Times New Roman"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =6551
                    Width =4649
                    Height =450
                    Name ="Box52"
                End
                Begin Line
                    Left =1649
                    Top =6551
                    Width =0
                    Height =450
                    Name ="Line53"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =6611
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text54"
                    Caption ="12"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =226
                    Width =4588
                    Height =283
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Field58"
                    ControlSource ="=\"Event \" & [E_Num] & \": \" & [ET_Des]"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =174
                    Top =853
                    Width =1418
                    Height =227
                    FontWeight =700
                    TabIndex =5
                    BackColor =0
                    ForeColor =16777215
                    Name ="Field62"
                    ControlSource ="=DLookUp(\"[rHead1]\",\"Miscellaneous\")"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =1705
                    Top =853
                    Width =2993
                    Height =227
                    FontWeight =700
                    TabIndex =6
                    BackColor =0
                    ForeColor =16777215
                    Name ="Field64"
                    ControlSource ="=DLookUp(\"[rHead2]\",\"Miscellaneous\")"

                End
                Begin Rectangle
                    BackStyle =0
                    Left =118
                    Top =1590
                    Width =4649
                    Height =465
                    Name ="Box65"
                End
                Begin Line
                    Left =1649
                    Top =1590
                    Width =0
                    Height =450
                    Name ="Line66"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =286
                    Top =1650
                    Width =1200
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Text67"
                    Caption ="1"
                    FontName ="Times New Roman"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter2"
        End
    End
End
