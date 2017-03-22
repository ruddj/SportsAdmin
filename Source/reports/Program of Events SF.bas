Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =50
    GridY =50
    Width =8221
    DatasheetFontHeight =10
    ItemSuffix =20
    Left =330
    Top =270
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf0164870efe5e140
    End
    RecordSource ="Program of Events SF"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x3702000037020000370200006e04000000000000fa0900001b01000000000000 ,
        0x040000007100000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
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
            ControlSource ="Lane"
        End
        Begin BreakLevel
            ControlSource ="Surname"
        End
        Begin BreakLevel
            ControlSource ="Gname"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =237
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Left =30
                    Width =2226
                    Height =193
                    FontSize =7
                    Name ="Text6"
                    ControlSource ="=UCase([Surname]) & \", \" & [Gname] & \" (\" & [H_Code] & \")\""

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =2278
                    Width =261
                    Height =193
                    FontSize =7
                    TabIndex =1
                    Name ="Text7"
                    ControlSource ="Lane"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =56
                    Width =5735
                    Name ="Line14"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
    End
End
