Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    GridY =10
    Width =5669
    ItemSuffix =3
    Left =1050
    Top =210
    Right =7365
    Bottom =7590
    RecSrcDt = Begin
        0xa51a1eba11cde140
    End
    RecordSource ="Sex Sub"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =288
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Width =1266
                    Height =225
                    Name ="Field0"
                    ControlSource ="Sex"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =1360
                    Width =1371
                    Height =225
                    TabIndex =1
                    Name ="Field2"
                    ControlSource ="Sex Sub"
                    FontName ="Tahoma"

                End
            End
        End
    End
End
