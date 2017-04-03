Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =5555
    ItemSuffix =15
    Left =1050
    Top =210
    Right =7320
    Bottom =6465
    RecSrcDt = Begin
        0x5ed8e9b911cde140
    End
    RecordSource ="SELECT DISTINCTROW [Lane Sub].Lane, [Lane Sub].Lane_Sub FROM [Lane Sub] ORDER BY"
        " [Lane Sub].Lane;"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
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
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            Height =288
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =1299
                    Height =227
                    Name ="Field2"
                    ControlSource ="Lane"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1432
                    Width =1254
                    Height =227
                    TabIndex =1
                    Name ="Field13"
                    ControlSource ="Lane_Sub"
                    FontName ="Tahoma"

                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter2"
        End
    End
End
