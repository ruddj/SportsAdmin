Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10703
    DatasheetFontHeight =10
    ItemSuffix =47
    Left =195
    Top =330
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xe9ab34e08f2ee240
    End
    RecordSource ="Program of Events"
    Caption ="Program of Events - Summary"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x370200003702000037020000d002000000000000cf2900000000000001000000 ,
        0x010000006801000000000000a20700000100000001000000
    End
    FilterOnLoad =255
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin BreakLevel
            KeepTogether =1
            ControlSource ="E_Number"
        End
        Begin BreakLevel
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="Sex"
        End
        Begin BreakLevel
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="F_Lev"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Heat"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1264
            Name ="PageHeader"
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Top =1020
                    Width =10673
                    Height =244
                    BackColor =12632256
                    Name ="Box36"
                End
                Begin TextBox
                    TextFontFamily =18
                    Width =10701
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
                    Left =56
                    Top =453
                    Width =10641
                    Name ="Line74"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    TextFontFamily =18
                    Left =56
                    Top =510
                    Width =9120
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text102"
                    Caption ="Program of Events"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =45
                    Top =1020
                    Width =660
                    Height =225
                    FontWeight =700
                    Name ="Label12"
                    Caption ="Event #"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =9195
                    Top =1020
                    Width =1290
                    Height =225
                    FontWeight =700
                    Name ="TimeLab"
                    Caption ="Event Time"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =795
                    Top =1020
                    Width =1695
                    Height =225
                    FontWeight =700
                    Name ="Label34"
                    Caption ="Event Description"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6885
                    Top =1020
                    Width =2130
                    Height =225
                    FontWeight =700
                    Name ="Label37"
                    Caption ="Record"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =355
            BreakLevel =5
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    Left =780
                    Top =60
                    Width =6006
                    Name ="Text0"
                    ControlSource ="=([ET_Des] & \"  /  \" & [Sex Sub] & \"  /  Age \" & [Age] & \"  /  \" & [F_Lev_"
                        "Sub] & \"  /  Heat \" & [Heat])"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =56
                    Top =59
                    Width =591
                    FontWeight =700
                    TabIndex =1
                    Name ="Text5"
                    ControlSource ="E_Number"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =9127
                    Top =59
                    Width =1476
                    FontWeight =700
                    TabIndex =2
                    Name ="E_Time"
                    ControlSource ="E_Time2"
                    Format ="d/mm/yy h:nn ampm"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    BackStyle =0
                    Left =6885
                    Top =59
                    Width =1446
                    TabIndex =3
                    Name ="Text38"
                    ControlSource ="=IIf(nz([Record])<>\"\",[Record] & \" \" & [Units],\"\")"

                End
                Begin Line
                    BorderWidth =1
                    Left =6870
                    Width =0
                    Height =344
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =1
                    Left =9070
                    Width =0
                    Height =344
                    Name ="Line40"
                End
                Begin Line
                    BorderWidth =1
                    Left =705
                    Width =0
                    Height =344
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =344
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =1
                    Left =10650
                    Width =0
                    Height =344
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =1
                    Top =340
                    Width =10658
                    Name ="Line45"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    Left =8385
                    Top =60
                    Width =666
                    TabIndex =4
                    Name ="Text46"
                    ControlSource ="H_Code"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =15
            Name ="Detail"
            Begin
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =56
                    Width =4025
                    Name ="Line14"
                End
            End
        End
        Begin PageFooter
            Height =390
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Width =1131
                    Height =390
                    FontSize =6
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
                    Width =9426
                    Height =390
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    Name ="Field88"
                    ControlSource ="=DLookUp(\"[CarnivalFooter]\",\"Miscellaneous\")"
                    FontName ="Times New Roman"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =2
                    BorderLineStyle =3
                    Width =10596
                    Name ="Line87"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =56
            Name ="ReportFooter"
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Top =15
                    Width =10673
                    Height =34
                    BackColor =12632256
                    Name ="Box44"
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

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
  If IsNull(Me!E_Time) Then
    Me!E_Time.visible = False
    'Me!TimeLab.Visible = False
  Else
    Me!E_Time.visible = True
    'Me!TimeLab.Visible = True
  End If
End Sub

Private Sub Report_NoData(Cancel As Integer)
  Dim Response As Integer
  Response = MsgBox("There is no data to display for the report: " & Me.Caption, vbInformation)
  Cancel = True

End Sub
