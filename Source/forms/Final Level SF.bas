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
    Bottom =6225
    RecSrcDt = Begin
        0x06fcb8b911cde140
    End
    RecordSource ="SELECT DISTINCTROW [Final Level Sub].F_Lev, [Final Level Sub].F_Lev_Sub FROM [Fi"
        "nal Level Sub] ORDER BY [Final Level Sub].F_Lev;"
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
                    Width =729
                    Height =227
                    Name ="Field2"
                    ControlSource ="F_Lev"
                    FontName ="Arial"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =793
                    Width =2229
                    Height =227
                    TabIndex =1
                    Name ="Field13"
                    ControlSource ="F_Lev_Sub"
                    FontName ="Arial"

                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter2"
        End
    End
End
