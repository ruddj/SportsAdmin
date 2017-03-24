Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =15953
    ItemSuffix =27
    Left =7470
    Top =3435
    Right =11520
    Bottom =7590
    HelpContextId =120
    OrderBy ="[Import Competitors SF].H_Code"
    RecSrcDt = Begin
        0xe3ed7db97706e240
    End
    RecordSource ="SELECT DISTINCTROW [Import Competitors].Sname, [Import Competitors].Gname, [Impo"
        "rt Competitors].H_Code, [Import Competitors].Age, [Import Competitors].DOB, [Imp"
        "ort Competitors].Sex, [Import Competitors].PIN FROM [Import Competitors] ORDER B"
        "Y [Import Competitors].Sname, [Import Competitors].Gname, [Import Competitors].H"
        "_Code, [Import Competitors].Age, [Import Competitors].DOB, [Import Competitors]."
        "Sex;"
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
            Height =270
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4515
                    Top =15
                    Width =735
                    Height =240
                    TabIndex =3
                    Name ="Sex"
                    ControlSource ="Sex"
                    Format =">"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3540
                    Top =15
                    Width =945
                    Height =240
                    TabIndex =2
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    Format =">"
                    StatusBarText ="House Code"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5280
                    Top =15
                    Width =660
                    Height =240
                    TabIndex =4
                    Name ="Age"
                    ControlSource ="Age"
                    Format =">"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =75
                    Top =15
                    Width =1815
                    Height =240
                    Name ="Gname"
                    ControlSource ="Gname"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =1920
                    Top =15
                    Width =1590
                    Height =240
                    TabIndex =1
                    Name ="Sname"
                    ControlSource ="Sname"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5970
                    Top =15
                    Width =945
                    Height =240
                    TabIndex =5
                    Name ="DOB"
                    ControlSource ="DOB"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =6945
                    Top =15
                    Width =945
                    Height =240
                    TabIndex =6
                    Name ="Text26"
                    ControlSource ="PIN"
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
