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
    Width =10671
    DatasheetFontHeight =10
    ItemSuffix =35
    Left =3075
    Top =240
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    Filter ="([Status]=0 OR [Status]=1 OR [Status]=2 OR [Status]=3)"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x94fcd77b6dd7e240
    End
    RecordSource ="Program of Events"
    Caption ="Program of Events"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x370200003702000037020000d002000000000000fc2800001b01000000000000 ,
        0x010000006801000000000000a20700000100000001000000
    End
    FilterOnLoad =0
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
            GroupHeader = NotDefault
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="F_Lev"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Heat"
        End
        Begin PageHeader
            Height =963
            Name ="PageHeader"
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
                    Left =56
                    Top =453
                    Width =10596
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
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =15
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader1"
            Begin
                Begin Line
                    BorderWidth =1
                    Width =10651
                    Name ="Line17"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =907
            BreakLevel =5
            Name ="GroupHeader0"
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =10651
                    Height =278
                    BackColor =15395562
                    Name ="Box10"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    Left =56
                    Width =7101
                    FontWeight =700
                    Name ="Text0"
                    ControlSource ="=([ET_Des] & \"  /  \" & [Sex Sub] & \"  /  Age \" & [Age] & \"  /  \" & [F_Lev_"
                        "Sub] & \"  /  Heat \" & [Heat])"

                End
                Begin Label
                    TextFontFamily =34
                    Top =285
                    Width =1140
                    Height =225
                    FontWeight =700
                    Name ="Label15"
                    Caption ="Competitor"
                End
                Begin Label
                    TextFontFamily =34
                    Left =2040
                    Top =283
                    Width =540
                    Height =225
                    FontWeight =700
                    Name ="Label16"
                    Caption ="Lane"
                End
                Begin Subform
                    OldBorderStyle =0
                    Top =567
                    Width =10671
                    Height =282
                    TabIndex =1
                    Name ="Program of Events SF"
                    SourceObject ="Report.Program of Events SF"
                    LinkChildFields ="F_Lev;Heat;E_Code"
                    LinkMasterFields ="F_Lev;Heat;E_Code"
                    EventProcPrefix ="Program_of_Events_SF"

                End
                Begin Label
                    TextFontFamily =34
                    Left =2664
                    Top =285
                    Width =1140
                    Height =225
                    FontWeight =700
                    Name ="Label26"
                    Caption ="Competitor"
                End
                Begin Label
                    TextFontFamily =34
                    Left =4704
                    Top =283
                    Width =540
                    Height =225
                    FontWeight =700
                    Name ="Label27"
                    Caption ="Lane"
                End
                Begin Label
                    TextFontFamily =34
                    Left =5329
                    Top =285
                    Width =1140
                    Height =225
                    FontWeight =700
                    Name ="Label28"
                    Caption ="Competitor"
                End
                Begin Label
                    TextFontFamily =34
                    Left =7369
                    Top =283
                    Width =540
                    Height =225
                    FontWeight =700
                    Name ="Label29"
                    Caption ="Lane"
                End
                Begin Label
                    TextFontFamily =34
                    Left =7993
                    Top =285
                    Width =1140
                    Height =225
                    FontWeight =700
                    Name ="Label30"
                    Caption ="Competitor"
                End
                Begin Label
                    TextFontFamily =34
                    Left =10033
                    Top =283
                    Width =540
                    Height =225
                    FontWeight =700
                    Name ="Label31"
                    Caption ="Lane"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    Left =10147
                    Width =456
                    FontWeight =700
                    TabIndex =2
                    Name ="Text5"
                    ControlSource ="E_Number"

                End
                Begin Label
                    TextAlign =3
                    Left =9420
                    Width =675
                    Height =225
                    Name ="Label12"
                    Caption ="Event #:"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    Left =7927
                    Width =1491
                    FontWeight =700
                    TabIndex =3
                    Name ="E_Time"
                    ControlSource ="E_Time2"

                End
                Begin Label
                    TextAlign =3
                    Left =7200
                    Width =675
                    Height =225
                    Name ="TimeLab"
                    Caption ="Time:"
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
    Me!TimeLab.visible = False
  Else
    Me!E_Time.visible = True
    Me!TimeLab.visible = True
  End If
End Sub

Private Sub Report_NoData(Cancel As Integer)
  Dim Response As Integer
  
  Response = MsgBox("There is no data to display for the report: " & Me.Caption, vbInformation)
  Cancel = True

End Sub
