Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =132
    Left =1515
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x7a4145ba11cde140
    End
    RecordSource ="SELECT DISTINCTROW House.H_NAme, [Surname] & \", \" & [Gname] AS Fullname, Compe"
        "titors.PIN, [Sex Sub].[Sex Sub], Competitors.Age FROM House INNER JOIN (Competit"
        "ors INNER JOIN [Sex Sub] ON Competitors.Sex = [Sex Sub].Sex) ON House.H_Code = C"
        "ompetitors.H_Code WHERE ((House.Flag=True)) ORDER BY [Surname] & \", \" & [Gname"
        "];"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x3702000037020000370200003702000000000000a02900008c01000001000000 ,
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
        Begin Chart
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="H_NAme"
        End
        Begin BreakLevel
            ControlSource ="Fullname"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =1
            ControlSource ="PIN"
        End
        Begin PageHeader
            Height =1946
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
                    Top =453
                    Width =10596
                    Name ="Line74"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =56
                    Top =1584
                    Width =1365
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text107"
                    Caption ="Name"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =4534
                    Top =1586
                    Width =630
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text108"
                    Caption ="Age"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Top =1133
                    Width =8226
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="Field115"
                    ControlSource ="=[H_NAme]"

                End
                Begin Label
                    TextFontFamily =18
                    Top =630
                    Width =3945
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text129"
                    Caption ="Competitor List"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =6121
                    Top =1586
                    Width =630
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text130"
                    Caption ="DOB"
                    FontName ="times New Roman"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="GroupHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =113
            BreakLevel =2
            Name ="GroupHeader2"
            Begin
                Begin Line
                    BorderWidth =2
                    Left =56
                    Top =56
                    Width =10437
                    Name ="Line112"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =396
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =116
                    Width =4086
                    Height =285
                    FontSize =10
                    Name ="Field98"
                    ControlSource ="Fullname"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4365
                    Width =1011
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Field100"
                    ControlSource ="Age"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =56
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =446
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9524
                    Top =56
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =56
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
                    Top =56
                    Width =10596
                    Name ="Line87"
                End
            End
        End
    End
End
