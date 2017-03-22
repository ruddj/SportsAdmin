Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =138
    Left =510
    Top =315
    ShortcutMenuBar ="cmdReportRightClick"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xc8bc2cb1efe5e140
    End
    RecordSource ="SELECT DISTINCTROW EventType.ET_Des, House.H_NAme, [Surname] & \", \" & [Gname] "
        "AS Fullname, Competitors.PIN, [Sex Sub].[Sex Sub], Events.Age, CompEvents.Points"
        ", [Result] & ' ' & [Units] AS fResult, IIf([Place]=0,'-',Str([Place])) AS PlaceS"
        ", [Final Level Sub].F_Lev_Sub, [Final Level Sub].F_Lev, IIf([Place]=0,1E+31,[Pla"
        "ce]) AS PlaceN, House.H_Code, CompEvents.Place FROM House RIGHT JOIN (EventType "
        "LEFT JOIN ((Competitors LEFT JOIN [Sex Sub] ON Competitors.Sex = [Sex Sub].Sex) "
        "RIGHT JOIN ((Events LEFT JOIN CompEvents ON Events.E_Code = CompEvents.E_Code) L"
        "EFT JOIN [Final Level Sub] ON CompEvents.F_Lev = [Final Level Sub].F_Lev) ON Com"
        "petitors.PIN = CompEvents.PIN) ON EventType.ET_Code = Events.ET_Code) ON House.H"
        "_Code = Competitors.H_Code WHERE (((House.Include)=True) AND ((House.Flag)=True)"
        " AND ((EventType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.Include"
        ")=True) AND ((CompEvents.Place) Is Not Null));"
    OnOpen ="[Event Procedure]"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200003702000037020000d002000000000000a02900008c01000001000000 ,
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
        Begin Chart
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="Sex Sub"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="F_Lev"
        End
        Begin BreakLevel
            ControlSource ="PlaceN"
        End
        Begin BreakLevel
            ControlSource ="Fullname"
        End
        Begin PageHeader
            Height =1077
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10386
                    Height =450
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
                    Top =453
                    Width =10596
                    Name ="Line74"
                End
                Begin Label
                    TextFontFamily =18
                    Top =630
                    Width =3945
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text129"
                    Caption ="Event Results"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextAlign =3
                    Left =9977
                    Top =113
                    Width =621
                    TabIndex =1
                    Name ="PageNo"
                    ControlSource ="=\"Pg \" & [Page]"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =963
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =5101
                    Top =509
                    Width =2280
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text122"
                    Caption ="Event Description"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =56
                    Top =507
                    Width =1365
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text107"
                    Caption ="Name"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =4144
                    Top =509
                    Width =630
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text108"
                    Caption ="Age"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =9751
                    Top =510
                    Width =825
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Points"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Top =56
                    Width =8226
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Field115"
                    ControlSource ="=[ET_Des] & ' - ' & [Sex Sub] & '  ' & [Age] & '  ' & [F_Lev_Sub]"

                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =7654
                    Top =510
                    Width =675
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text124"
                    Caption ="Place"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =8451
                    Top =510
                    Width =1305
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text126"
                    Caption ="Result"
                    FontName ="times New Roman"
                End
                Begin Line
                    BorderWidth =2
                    Left =56
                    Top =907
                    Width =10527
                    Height =15
                    Name ="Line132"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =3288
                    Top =510
                    Width =690
                    Height =345
                    FontSize =13
                    FontWeight =700
                    Name ="Text137"
                    Caption ="Team"
                    FontName ="times New Roman"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =396
            OnFormat ="[Event Procedure]"
            OnRetreat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7596
                    Width =741
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="Field125"
                    ControlSource ="PlaceS"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =116
                    Width =3111
                    Height =285
                    FontSize =10
                    Name ="Field98"
                    ControlSource ="Fullname"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4155
                    Width =831
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Field100"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9993
                    Width =531
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Field106"
                    ControlSource ="Points"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4977
                    Width =2616
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Field123"
                    ControlSource ="ET_Des"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =8333
                    Width =1461
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="Field127"
                    ControlSource ="fResult"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =56
                    Top =340
                    Width =10527
                    Height =15
                    Name ="Line112"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3240
                    Width =966
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="Field136"
                    ControlSource ="H_Code"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =226
            BreakLevel =3
            OnRetreat ="[Event Procedure]"
            Name ="GroupFooter0"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =446
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9524
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
                    Width =9471
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

Option Compare Database   'Use database order for string comparisons

Option Explicit

Dim DisplayRecords As Variant
Dim NumberToDisplay As Variant

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    'MsgBox ("Display Records = " & Str(DisplayRecords))
    'Stop
    If DisplayRecords >= NumberToDisplay Then
        Cancel = True
    End If

    DisplayRecords = DisplayRecords + 1

End Sub

Private Sub Detail1_Retreat()

    DisplayRecords = 0

End Sub

Private Sub GroupFooter0_Retreat()

    DisplayRecords = 0

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)

    DisplayRecords = 0

End Sub

Private Sub Report_Open(Cancel As Integer)

On Error Resume Next

    DisplayRecords = 0
    'NumberToDisplay = DLookup("[CompetitorPlaces]", "MiscellaneousLocal")
    NumberToDisplay = DLookup("[NumberOfRecords]", "Misc-Statistics")
    If IsNull(NumberToDisplay) Then
      NumberToDisplay = 1
    End If


End Sub
