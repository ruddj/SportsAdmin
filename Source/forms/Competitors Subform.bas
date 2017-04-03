Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =0
    GridX =20
    GridY =20
    Width =5102
    ItemSuffix =13
    Left =2535
    Top =990
    Right =9690
    Bottom =4965
    HelpContextId =70
    RecSrcDt = Begin
        0x75173b042dc7e140
    End
    RecordSource ="CompetitorsSubform"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =72
            BackColor =-2147483633
            Name ="FormHeader0"
        End
        Begin Section
            Height =330
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    Left =72
                    Width =1963
                    Height =227
                    BackColor =-2147483633
                    Name ="Field0"
                    ControlSource ="ET_Des"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    Left =2006
                    Width =459
                    Height =227
                    TabIndex =1
                    BackColor =-2147483633
                    Name ="Field3"
                    ControlSource ="F_Lev"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    Left =2922
                    Width =411
                    Height =227
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483633
                    ForeColor =255
                    Name ="Field5"
                    ControlSource ="Place"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    Left =3330
                    Width =984
                    Height =227
                    FontSize =7
                    TabIndex =4
                    BackColor =-2147483633
                    Name ="Field10"
                    ControlSource ="res"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    BackStyle =0
                    Left =4320
                    Width =486
                    Height =227
                    FontWeight =700
                    TabIndex =3
                    BackColor =-2147483633
                    ForeColor =255
                    Name ="PtsFld"
                    ControlSource ="Points"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    Left =2488
                    Width =399
                    Height =227
                    TabIndex =5
                    BackColor =-2147483633
                    Name ="Field11"
                    ControlSource ="Heat"
                    FontName ="Tahoma"

                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =75
                    Top =285
                    Width =4764
                    Name ="Line12"
                End
            End
        End
        Begin FormFooter
            Height =453
            BackColor =-2147483633
            Name ="FormFooter1"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    Left =3798
                    Top =113
                    Width =794
                    Height =227
                    FontWeight =700
                    BorderColor =12632256
                    Name ="TotFld"
                    ControlSource ="=Sum([Points])"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2777
                            Top =113
                            Width =915
                            Height =240
                            BackColor =8454143
                            Name ="Text8"
                            Caption ="Total Points"
                        End
                    End
                End
            End
        End
    End
End
