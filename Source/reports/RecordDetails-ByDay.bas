Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10714
    ItemSuffix =143
    Left =1500
    Top =240
    OnNoData ="[Event Procedure]"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x952adbb3d0dae140
    End
    RecordSource ="SELECT DISTINCTROW EventType.ET_Des, EventType.Units, Events.Sex, Events.Age, Ho"
        "use.H_NAme, House.H_Code, IIf([Order]=\"ASC\",[nResult],(1/[nResult])) AS BestRe"
        "sult, Records.Result, [Gname] & \" \" & UCase([Surname]) AS FullName, [Sex Sub]."
        "[Sex Sub], EventType.Include, EventType.Flag, House.Flag, Records.Date FROM (Eve"
        "ntType LEFT JOIN Units ON EventType.Units = Units.DisplayUnit) RIGHT JOIN ((Even"
        "ts LEFT JOIN [Sex Sub] ON Events.Sex = [Sex Sub].Sex) RIGHT JOIN (House RIGHT JO"
        "IN Records ON House.H_Code = Records.H_Code) ON Events.E_Code = Records.E_Code) "
        "ON EventType.ET_Code = Events.ET_Code WHERE (((EventType.Include)=Yes) AND ((Eve"
        "ntType.Flag)=Yes) AND ((House.Flag)=Yes) AND ((Records.Date)=DLookUp(\"[RecordDa"
        "te]\",\"Misc-Statistics\")));"
    Caption ="Records set on the Specified Day"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200003702000037020000d002000000000000da2900001801000001000000 ,
        0x010000006801000000000000a10700000100000000000000
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
        Begin Chart
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =1
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Sex"
        End
        Begin BreakLevel
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="BestResult"
        End
        Begin PageHeader
            Height =971
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10656
                    Height =405
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
                    Top =453
                    Width =10596
                    Name ="Line74"
                End
                Begin TextBox
                    TextFontFamily =18
                    Top =566
                    Width =10656
                    Height =390
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    Name ="Field141"
                    ControlSource ="=\"RECORDS SET ON \" & Format(DLookUp(\"[RecordDate]\",\"Misc-Statistics\"),\"Lo"
                        "ng Date\")"
                    FontName ="times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =738
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =56
                    Top =56
                    Width =8346
                    Height =285
                    FontSize =11
                    FontWeight =700
                    Name ="Field118"
                    ControlSource ="=\"EVENT: \" & [ET_Des]"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextFontFamily =34
                    Left =283
                    Top =453
                    Width =960
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text126"
                    Caption ="DIVISION"
                End
                Begin Label
                    TextFontFamily =34
                    Left =2211
                    Top =453
                    Width =660
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text127"
                    Caption ="NAME"
                End
                Begin Label
                    TextFontFamily =34
                    Left =5612
                    Top =453
                    Width =645
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text128"
                    Caption ="TEAM"
                End
                Begin Label
                    TextFontFamily =34
                    Left =7483
                    Top =453
                    Width =885
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text129"
                    Caption ="RESULT"
                End
                Begin Label
                    TextFontFamily =34
                    Left =9411
                    Top =453
                    Width =615
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text130"
                    Caption ="DATE"
                End
                Begin Line
                    BorderWidth =2
                    Left =56
                    Top =396
                    Width =10545
                    Name ="Line140"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader1"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =280
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =2199
                    Width =3231
                    Height =225
                    FontSize =9
                    Name ="Field98"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5604
                    Width =1716
                    Height =225
                    FontSize =9
                    TabIndex =1
                    Name ="Field117"
                    ControlSource ="H_NAme"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =285
                    Width =801
                    Height =225
                    FontSize =9
                    TabIndex =2
                    Name ="Field119"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =7486
                    Width =846
                    Height =225
                    FontSize =9
                    TabIndex =3
                    Name ="Field121"
                    ControlSource ="Result"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =8400
                    Width =906
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="Field122"
                    ControlSource ="Units"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =1140
                    Width =906
                    Height =225
                    FontSize =9
                    TabIndex =5
                    Name ="Field125"
                    ControlSource ="Sex Sub"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =9424
                    Width =846
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="Field131"
                    ControlSource ="Date"

                End
                Begin Rectangle
                    BackStyle =0
                    Left =165
                    Width =1927
                    Height =280
                    Name ="Box135"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =2092
                    Width =3397
                    Height =280
                    Name ="Box136"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =5484
                    Width =1882
                    Height =280
                    Name ="Box137"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =7370
                    Width =1927
                    Height =280
                    Name ="Box138"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =9302
                    Width =1072
                    Height =280
                    Name ="Box139"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =113
            BreakLevel =1
            Name ="GroupFooter2"
        End
        Begin PageFooter
            Height =446
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Top =56
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Top =56
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
                    Top =56
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
Option Compare Database
Option Explicit

Private Sub Report_NoData(Cancel As Integer)

  MsgBox ("There is no data to display for the report: " & Me.Caption)
  Cancel = True
  
End Sub
