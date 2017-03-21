Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    GridX =50
    GridY =50
    Width =6988
    ItemSuffix =5
    Left =1785
    Top =1320
    Right =11520
    Bottom =7590
    HelpContextId =40
    RecSrcDt = Begin
        0x5fa1726d32cde140
    End
    RecordSource ="House Points-Extra"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =288
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =86
                    Top =5
                    Width =465
                    Height =225
                    Name ="Text5"
                    Caption ="# Pts"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    Left =875
                    Width =4020
                    Height =225
                    Name ="Text6"
                    Caption ="Reason for allocation / deallocation of points"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =346
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =72
                    Width =726
                    Height =285
                    Name ="NumPts"
                    ControlSource ="NumPts"
                    ControlTipText ="Number of extra points to add or subtract."

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =864
                    Width =4641
                    Height =285
                    TabIndex =1
                    Name ="Reason"
                    ControlSource ="Reason"
                    ControlTipText ="Optional: The reason for the extra points."

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6292
                    Width =696
                    Height =225
                    TabIndex =2
                    Name ="H_ID"
                    ControlSource ="H_ID"

                End
            End
        End
        Begin FormFooter
            Height =144
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
