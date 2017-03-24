Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =15953
    ItemSuffix =25
    Left =7530
    Top =3150
    Right =11520
    Bottom =7590
    HelpContextId =150
    Filter ="(([Import Data SF].S_Name=\"Biggs\"))"
    OrderBy ="Age"
    RecSrcDt = Begin
        0x273cc99d7706e240
    End
    RecordSource ="SELECT DISTINCTROW ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, Impo"
        "rtData.Age, ImportData.Sex, ImportData.HE_Code, ImportData.ET_Des, ImportData.He"
        "at, ImportData.Competitor, ImportData.Memo FROM ImportData ORDER BY ImportData.S"
        "_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age;"
    Caption ="ImportData"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin OptionButton
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =187
            Height =187
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =255
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
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =875
                    Width =570
                    Height =230
                    TabIndex =1
                    Name ="Sex"
                    ControlSource ="Sex"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =2154
                    Width =2010
                    Height =230
                    TabIndex =3
                    Name ="ET_Des"
                    ControlSource ="ET_Des"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4680
                    Width =285
                    Height =230
                    TabIndex =5
                    Name ="Competitor"
                    ControlSource ="Competitor"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =720
                    Height =230
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    Format =">"
                    StatusBarText ="House Code"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1510
                    Width =570
                    Height =230
                    TabIndex =2
                    Name ="Age"
                    ControlSource ="Age"
                    Format =">"
                    FontName ="Arial"

                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4240
                    Width =360
                    Height =230
                    TabIndex =4
                    Name ="Heat"
                    ControlSource ="Heat"
                    Format =">"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =5040
                    Width =1290
                    Height =230
                    TabIndex =6
                    Name ="G_Name"
                    ControlSource ="G_Name"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =6404
                    Width =1380
                    Height =230
                    TabIndex =7
                    Name ="S_Name"
                    ControlSource ="S_Name"
                    FontName ="Arial"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter2"
        End
    End
End
