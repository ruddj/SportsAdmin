Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10977
    ItemSuffix =150
    Left =1500
    Top =240
    OnNoData ="[Event Procedure]"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x0fb69d24ecdae140
    End
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8d010000370200008d0100005303000000000000e12a00000000000001000000 ,
        0x010000007100000000000000a20700000100000000000000
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
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Chart
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            KeepTogether =1
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="ET_Code"
        End
        Begin PageHeader
            Height =907
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10941
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
                    Width =10881
                    Name ="Line74"
                End
                Begin Label
                    TextFontFamily =18
                    Left =113
                    Top =510
                    Width =2925
                    Height =330
                    FontSize =15
                    FontWeight =700
                    Name ="Text102"
                    Caption ="RECORD HOLDERS"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =3685
                    Top =567
                    TabIndex =1
                    Name ="Female"
                    ControlSource ="=\"F\""

                End
                Begin TextBox
                    Visible = NotDefault
                    Left =5442
                    Top =567
                    TabIndex =2
                    Name ="Male"
                    ControlSource ="=\"M\""

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =453
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin Subform
                    Left =56
                    Top =56
                    Width =5383
                    Height =309
                    Name ="Embedded141"
                    SourceObject ="Report.RecordDetailsMini-SF"
                    LinkChildFields ="Sex"
                    LinkMasterFields ="Female"

                End
                Begin Subform
                    Left =5612
                    Top =56
                    Width =5353
                    Height =279
                    TabIndex =1
                    Name ="Embedded148"
                    SourceObject ="Report.RecordDetailsMini-SF"
                    LinkChildFields ="Sex"
                    LinkMasterFields ="Male"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =0
            Name ="Detail1"
        End
        Begin PageFooter
            Height =446
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9751
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
                    Width =9591
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
                    Width =10896
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
