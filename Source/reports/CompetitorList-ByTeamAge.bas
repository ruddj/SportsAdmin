Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10725
    ItemSuffix =152
    Left =855
    Top =645
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xef2b10c2f317e240
    End
    RecordSource ="Competitor List"
    OnOpen ="[Event Procedure]"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x3702000037020000370200009702000000000000e52900006301000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    RibbonName ="SportPrint"
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
            ControlSource ="AAge"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Sex"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="H_Code"
        End
        Begin BreakLevel
            ControlSource ="FullName"
        End
        Begin PageHeader
            Height =1296
            Name ="PageHeader0"
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Top =750
                    Width =10680
                    Height =411
                    BackColor =12632256
                    Name ="Box151"
                End
                Begin TextBox
                    TextFontFamily =18
                    Width =10041
                    Height =450
                    ColumnOrder =0
                    FontSize =16
                    FontWeight =700
                    Name ="CarnivalTitle"
                    ControlSource ="=DLookUp(\"[CarnivalTitle]\",\"Miscellaneous\")"
                    FontName ="times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =3
                    BorderLineStyle =3
                    Left =56
                    Top =566
                    Width =10596
                    Name ="Line74"
                End
                Begin TextBox
                    TextFontFamily =18
                    Left =1756
                    Top =793
                    Width =2497
                    Height =347
                    ColumnOrder =1
                    FontSize =13
                    TabIndex =1
                    BackColor =12632256
                    Name ="Field144"
                    ControlSource ="=[Sex Sub] & \" \" & [AAge]"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            BackStyle =0
                            TextFontFamily =18
                            Top =793
                            Width =1755
                            Height =360
                            FontSize =13
                            FontWeight =700
                            Name ="Text145"
                            Caption ="AGE GROUP:"
                            FontName ="times New Roman"
                        End
                    End
                End
                Begin Label
                    TextFontFamily =18
                    Left =4365
                    Top =793
                    Width =4155
                    Height =345
                    FontSize =13
                    FontWeight =700
                    BackColor =12632256
                    Name ="Text146"
                    Caption ="EVENT: _______________________"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Left =8565
                    Top =793
                    Width =2160
                    Height =345
                    FontSize =13
                    FontWeight =700
                    BackColor =12632256
                    Name ="Text147"
                    Caption ="REC.: __________"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =18
                    Left =9875
                    Width =832
                    Height =347
                    ColumnOrder =2
                    FontSize =9
                    TabIndex =2
                    Name ="Field149"
                    ControlSource ="=\"Page \" & [Page]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            ForceNewPage =1
            Height =0
            BreakLevel =1
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =731
            BreakLevel =2
            Name ="GroupHeader0"
            Begin
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =114
                    Top =432
                    Width =2700
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text39"
                    Caption ="COMPETITOR"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =716
                    Width =10647
                    Name ="Line96"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =3005
                    Top =432
                    Width =1425
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text138"
                    Caption ="DOB"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =4422
                    Top =432
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text139"
                    Caption ="1"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5329
                    Top =432
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text140"
                    Caption ="2"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =6236
                    Top =432
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text141"
                    Caption ="3"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =8731
                    Top =432
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text142"
                    Caption ="PLACE"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =9695
                    Top =432
                    Width =900
                    Height =270
                    FontSize =10
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Text143"
                    Caption ="POINTS"
                End
                Begin TextBox
                    TextFontFamily =34
                    Top =60
                    Width =4373
                    Height =325
                    FontSize =12
                    FontWeight =700
                    Name ="Field94"
                    ControlSource ="H_NAme"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =355
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =226
                    Width =2768
                    Height =265
                    FontSize =10
                    Name ="Field21"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3118
                    Width =1367
                    Height =265
                    FontSize =10
                    TabIndex =1
                    Name ="Field23"
                    ControlSource ="DOB"

                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3061
                    Width =0
                    Height =351
                    Name ="Line58"
                End
                Begin Line
                    BorderWidth =1
                    Left =5385
                    Width =0
                    Height =351
                    Name ="Line60"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =56
                    Width =0
                    Height =351
                    Name ="Line77"
                End
                Begin Line
                    BorderWidth =1
                    Left =6236
                    Width =0
                    Height =351
                    Name ="Line78"
                End
                Begin Line
                    BorderWidth =1
                    Left =10714
                    Width =0
                    Height =351
                    Name ="Line79"
                End
                Begin Line
                    BorderWidth =1
                    Left =4535
                    Width =0
                    Height =351
                    Name ="Line82"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =56
                    Top =340
                    Width =10647
                    Name ="Line86"
                End
                Begin Line
                    BorderWidth =1
                    Left =7086
                    Width =0
                    Height =351
                    Name ="Line97"
                End
                Begin Line
                    BorderWidth =1
                    Left =7936
                    Width =0
                    Height =351
                    Name ="Line98"
                End
                Begin Line
                    BorderWidth =1
                    Left =8786
                    Width =0
                    Height =351
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =1
                    Left =9751
                    Width =0
                    Height =351
                    Name ="Line100"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =737
            BreakLevel =2
            Name ="GroupFooter2"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3061
                    Width =0
                    Height =351
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =5385
                    Width =0
                    Height =351
                    Name ="Line107"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =56
                    Width =0
                    Height =351
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Left =6236
                    Width =0
                    Height =351
                    Name ="Line109"
                End
                Begin Line
                    BorderWidth =1
                    Left =10714
                    Width =0
                    Height =351
                    Name ="Line110"
                End
                Begin Line
                    BorderWidth =1
                    Left =4535
                    Width =0
                    Height =351
                    Name ="Line111"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =56
                    Top =340
                    Width =10647
                    Name ="Line112"
                End
                Begin Line
                    BorderWidth =1
                    Left =7086
                    Width =0
                    Height =351
                    Name ="Line113"
                End
                Begin Line
                    BorderWidth =1
                    Left =7936
                    Width =0
                    Height =351
                    Name ="Line114"
                End
                Begin Line
                    BorderWidth =1
                    Left =8786
                    Width =0
                    Height =351
                    Name ="Line115"
                End
                Begin Line
                    BorderWidth =1
                    Left =9751
                    Width =0
                    Height =351
                    Name ="Line116"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3061
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line127"
                End
                Begin Line
                    BorderWidth =1
                    Left =5385
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line128"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =56
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line129"
                End
                Begin Line
                    BorderWidth =1
                    Left =6236
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =10714
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line131"
                End
                Begin Line
                    BorderWidth =1
                    Left =4535
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line132"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =56
                    Top =680
                    Width =10647
                    Name ="Line133"
                End
                Begin Line
                    BorderWidth =1
                    Left =7086
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line134"
                End
                Begin Line
                    BorderWidth =1
                    Left =7936
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =1
                    Left =8786
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line136"
                End
                Begin Line
                    BorderWidth =1
                    Left =9751
                    Top =340
                    Width =0
                    Height =351
                    Name ="Line137"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            BreakLevel =1
            Name ="GroupFooter0"
        End
        Begin PageFooter
            Height =503
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Top =113
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =113
                    Width =9426
                    Height =390
                    FontSize =11
                    FontWeight =700
                    Name ="Field88"
                    ControlSource ="=DLookUp(\"[CarnivalFooter]\",\"Miscellaneous\")"
                    FontName ="Times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =2
                    BorderLineStyle =3
                    Top =113
                    Width =10596
                    Name ="Line87"
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons

Private Sub Report_Open(Cancel As Integer)
    
    Call UpdateEventCompetitorAge

    
End Sub
