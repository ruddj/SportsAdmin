﻿Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridX =50
    GridY =50
    Width =10675
    ItemSuffix =139
    Left =1605
    Top =2340
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xb95d1af2eee5e140
    End
    RecordSource ="Report-CompetitorResultsByEventTeam"
    OnOpen ="[Event Procedure]"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200003702000037020000d002000000000000b32900009001000001000000 ,
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
            GroupHeader = NotDefault
            ControlSource ="H_NAme"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Age"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="Sex Sub"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="Result"
        End
        Begin BreakLevel
            ControlSource ="Place"
        End
        Begin BreakLevel
            ControlSource ="Fullname"
        End
        Begin PageHeader
            Height =1133
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10656
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
                    Width =10080
                    Height =390
                    FontSize =15
                    FontWeight =700
                    Name ="Text129"
                    Caption ="Competitor Results Summary - Ordered by Result then Place then Name"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextAlign =3
                    Left =8955
                    Top =15
                    TabIndex =1
                    Name ="Field94"
                    ControlSource ="=\"Page \" & [Page]"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="GroupHeader3"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader4"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =813
            BreakLevel =2
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader1"
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =10656
                    Height =377
                    BackColor =12632256
                    Name ="Box137"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =3645
                    Top =450
                    Width =2445
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text122"
                    Caption ="Event Description"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =71
                    Top =451
                    Width =1365
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text107"
                    Caption ="Name"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =2846
                    Top =453
                    Width =630
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text108"
                    Caption ="Age"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =9940
                    Top =453
                    Width =705
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Pts"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    Left =15
                    Top =45
                    Width =8226
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="Field115"
                    ControlSource ="=[H_NAme] & ' - ' & [Sex Sub] & ' Results - Age: ' & [Age]"

                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =7895
                    Top =453
                    Width =675
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text124"
                    Caption ="Place"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =8689
                    Top =453
                    Width =1140
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text126"
                    Caption ="Result"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =6533
                    Top =453
                    Width =1245
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text132"
                    Caption ="Final"
                    FontName ="times New Roman"
                End
                Begin Line
                    BorderWidth =1
                    Left =56
                    Top =793
                    Width =10545
                    Name ="Line135"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =3
            BackColor =12632256
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =400
            OnFormat ="[Event Procedure]"
            OnRetreat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7823
                    Top =113
                    Width =741
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="Field125"
                    ControlSource ="Place"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =322
                    Top =115
                    Width =2421
                    Height =285
                    FontSize =10
                    Name ="Field98"
                    ControlSource ="Fullname"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =2777
                    Top =113
                    Width =786
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
                    Top =113
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
                    Left =3628
                    Top =113
                    Width =2901
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Field123"
                    ControlSource ="ET_Des"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =8557
                    Top =113
                    Width =1296
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="Field127"
                    ControlSource ="fResult"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =6521
                    Top =113
                    Width =1251
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="Field133"
                    ControlSource ="F_Lev_Sub"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =56
                    Top =56
                    Width =10545
                    Name ="DetailLine"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =1
                    TextFontFamily =34
                    Top =115
                    Width =261
                    Height =285
                    FontWeight =700
                    TabIndex =7
                    Name ="Text138"
                    ControlSource ="=1"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =73
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =58
                    Width =10620
                    Name ="Line134"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =113
            BreakLevel =2
            Name ="GroupFooter2"
        End
        Begin PageFooter
            Height =576
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

Dim NumberToDisplay As Variant

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)

    If DisplayRecords = 0 Then
        Me!DetailLine.visible = False
    Else
        Me!DetailLine.visible = True
    End If
    If DisplayRecords >= NumberToDisplay Then
        Cancel = True
    End If

    DisplayRecords = DisplayRecords + 1

End Sub

Private Sub Detail1_Retreat()

 '   DisplayRecords = 0

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    DisplayRecords = 0

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

 '   DisplayRecords = 0

End Sub

Private Sub Report_Open(Cancel As Integer)

    DisplayRecords = 0
    NumberToDisplay = DLookup("[NumberOfRecords]", "Misc-Statistics")
    If IsNull(NumberToDisplay) Then
      NumberToDisplay = 1
    End If

End Sub
