Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10620
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =600
    Top =210
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x12a5bd21f2e5e140
    End
    RecordSource ="SELECT Competitors.PIN, [Competitors that Competed in Flagged Events].PIN, Compe"
        "titors.Gname, Competitors.Surname, Competitors.Sex, Competitors.H_Code, Competit"
        "ors.Age FROM House RIGHT JOIN (Competitors LEFT JOIN [Competitors that Competed "
        "in Flagged Events] ON Competitors.PIN = [Competitors that Competed in Flagged Ev"
        "ents].PIN) ON House.H_Code = Competitors.H_Code WHERE (((House.Flag)=True) AND ("
        "(House.Include)=Yes) AND (([Competitors that Competed in Flagged Events].PIN) Is"
        " Null) AND ((Competitors.Gname)<>\"Team\"));"
    Caption ="Non-Participants"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd0020000d0020000d00200009602000000000000b01300002b01000000000000 ,
        0x020000006801000000000000a20700000100000001000000
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
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="Age"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="Sex"
        End
        Begin BreakLevel
            ControlSource ="Surname"
        End
        Begin BreakLevel
            ControlSource ="Gname"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =480
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =18
                    Top =60
                    Width =9900
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Competitors in the Database who did not participate in the flagged events"
                    FontName ="times New Roman"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =420
                    Width =10560
                    Name ="Line11"
                End
                Begin TextBox
                    TextAlign =3
                    Left =8880
                    Name ="Text13"
                    ControlSource ="=\"Page \" & [Page]"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Height =540
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =600
                    Top =120
                    Width =1080
                    Height =300
                    FontSize =11
                    FontWeight =700
                    Name ="Text1"
                    ControlSource ="Age"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Top =120
                            Width =600
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Label2"
                            Caption ="Age:"
                        End
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =2700
                    Top =120
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    Name ="Text5"
                    ControlSource ="Sex"

                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =1740
                            Top =120
                            Width =960
                            Height =285
                            FontSize =11
                            FontWeight =700
                            Name ="Label6"
                            Caption ="Gender:"
                        End
                    End
                End
                Begin Line
                    Top =480
                    Width =4800
                    Name ="Line7"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =4200
                    Top =120
                    Width =600
                    Height =285
                    FontSize =11
                    FontWeight =700
                    Name ="Cont"
                    Caption ="cont."
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    Left =60
                    Top =60
                    Width =3360
                    Name ="Text8"
                    ControlSource ="=UCase([Surname]) & \", \" & [Gname]"

                End
                Begin TextBox
                    Left =3480
                    Top =60
                    Width =1380
                    TabIndex =1
                    Name ="Text10"
                    ControlSource ="H_Code"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =510
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Top =60
                    Width =9471
                    Height =390
                    FontSize =11
                    FontWeight =700
                    Name ="Field88"
                    ControlSource ="=DLookUp(\"[CarnivalFooter]\",\"Miscellaneous\")"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9480
                    Top =60
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderWidth =2
                    BorderLineStyle =3
                    Top =60
                    Width =10596
                    Name ="Line87"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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

Dim SectionCount As Variant


Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
  SectionCount = 0
End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
  SectionCount = SectionCount + 1
  If SectionCount > 1 Then
    Me!Cont.visible = False
  Else
    Me!Cont.visible = False
  End If
  
End Sub

Private Sub PageHeader_Format(Cancel As Integer, FormatCount As Integer)
  SectionCount = 0
End Sub

Private Sub Report_Close()
  Call CreateReportList(Me.Name)
End Sub

Private Sub Report_NoData(Cancel As Integer)

  MsgBox ("There is no data to display for the report: " & Me.Caption)
  Cancel = True
  

End Sub
