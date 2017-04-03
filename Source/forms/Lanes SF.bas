Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =5555
    ItemSuffix =13
    Left =1965
    Top =1065
    Right =8235
    Bottom =7065
    HelpContextId =130
    RecSrcDt = Begin
        0xda2f3cf510cde140
    End
    RecordSource ="SELECT DISTINCTROW Lanes.Lane, Lanes.H_Code FROM Lanes ORDER BY Lanes.Lane;"
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
            Height =85
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =360
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =729
                    Height =287
                    Name ="Field2"
                    ControlSource ="Lane"
                    DefaultValue ="=DMax(\"[Lane]\",\"Lanes\")+2"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =2095
                    Left =865
                    Width =2665
                    Height =285
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Field11"
                    ControlSource ="H_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT House.H_ID, House.H_NAme, House.Include FROM House WHERE ((House.Include="
                        "Yes)) ORDER BY House.H_NAme;"
                    ColumnWidths ="0;2095;0"
                    FontName ="Tahoma"

                End
            End
        End
        Begin FormFooter
            Height =85
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
