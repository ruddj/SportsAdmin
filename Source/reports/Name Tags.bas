Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =4932
    ItemSuffix =15
    Left =630
    Top =165
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x7ae24de68df5e140
    End
    RecordSource ="SELECT DISTINCTROW House.Include, Competitors.PIN, Competitors.Gname, Competitor"
        "s.Surname, Competitors.Age, Competitors.Sex, House.H_Code, House.H_NAme FROM Hou"
        "se INNER JOIN Competitors ON House.H_Code = Competitors.H_Code WHERE (((House.In"
        "clude)=True) AND ((Competitors.Gname)<>\"Team\"));"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x37020000370200008f010000d002000000000000441300000f0f000001000000 ,
        0x020000005503000000000000a20700000100000000000000
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
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin BreakLevel
            ControlSource ="H_Code"
        End
        Begin BreakLevel
            ControlSource ="Surname"
        End
        Begin BreakLevel
            ControlSource ="Gname"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =3855
            Name ="Detail1"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =2
                    TextAlign =2
                    TextFontFamily =34
                    Left =56
                    Top =1125
                    Width =1119
                    Height =1941
                    FontSize =50
                    FontWeight =700
                    TabIndex =2
                    Name ="Field9"
                    ControlSource ="=Left([H_Code],1)"

                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =2
                    Left =56
                    Width =4867
                    Height =3056
                    Name ="Box0"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =18
                    Left =170
                    Top =56
                    Width =4632
                    Height =520
                    FontSize =18
                    FontWeight =700
                    Name ="Name"
                    ControlSource ="=[Gname] & ' ' & UCase([Surname])"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =170
                    Top =623
                    Width =840
                    Height =405
                    FontSize =16
                    FontWeight =700
                    Name ="Text4"
                    Caption ="AGE:"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =18
                    Left =1088
                    Top =629
                    Width =1572
                    Height =400
                    FontSize =16
                    TabIndex =1
                    Name ="Field7"
                    ControlSource ="Age"
                    FontName ="Times New Roman"

                End
                Begin Subform
                    CanGrow = NotDefault
                    Left =1267
                    Top =1187
                    Width =3613
                    Height =1809
                    TabIndex =3
                    Name ="Embedded11"
                    SourceObject ="Report.Name Tags SF"
                    LinkChildFields ="PIN"
                    LinkMasterFields ="PIN"

                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =2
                    Left =1190
                    Top =1125
                    Width =3742
                    Height =1936
                    Name ="Box13"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =2721
                    Top =680
                    Width =2142
                    Height =400
                    TabIndex =4
                    Name ="Text14"
                    ControlSource ="H_NAme"

                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter2"
        End
    End
End
