Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =135
    Left =2190
    Top =300
    ShortcutMenuBar ="cmdReportRightClick"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xd1daf824efe5e140
    End
    RecordSource ="SELECT DISTINCTROW EventType.Include, EventType.Flag, Events.Include, House.Incl"
        "ude, House.H_NAme, [Surname] & \", \" & [Gname] AS Fullname, Competitors.PIN, [S"
        "ex Sub].[Sex Sub], Events.Age, EventType.ET_Des, CompEvents.Points, [Result] & '"
        " ' & [Units] AS fResult, CompEvents.Place, [Final Level Sub].F_Lev_Sub, CompEven"
        "ts.F_Lev FROM House RIGHT JOIN (EventType RIGHT JOIN ((Competitors LEFT JOIN [Se"
        "x Sub] ON Competitors.Sex = [Sex Sub].Sex) LEFT JOIN ((Events RIGHT JOIN CompEve"
        "nts ON Events.E_Code = CompEvents.E_Code) LEFT JOIN [Final Level Sub] ON CompEve"
        "nts.F_Lev = [Final Level Sub].F_Lev) ON Competitors.PIN = CompEvents.PIN) ON Eve"
        "ntType.ET_Code = Events.ET_Code) ON House.H_Code = Competitors.H_Code WHERE (((E"
        "ventType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.Include)=True) "
        "AND ((House.Include)=True) AND ((House.Flag)=True)) ORDER BY [Surname] & \", \" "
        "& [Gname], CompEvents.Points DESC;"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x3702000037020000450200009602000000000000a02900003b01000001000000 ,
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
        Begin CommandButton
            TextFontFamily =2
            BorderLineStyle =0
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
            ControlSource ="H_NAme"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Sex Sub"
        End
        Begin BreakLevel
            ControlSource ="Fullname"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="PIN"
        End
        Begin BreakLevel
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="F_Lev"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader0"
        End
        Begin PageHeader
            Height =1947
            OnFormat ="[Event Procedure]"
            Name ="PageHeader0"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =3855
                    Top =1587
                    Width =1995
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text122"
                    Caption ="Description"
                    FontName ="times New Roman"
                End
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
                    TextAlign =1
                    TextFontFamily =18
                    Left =56
                    Top =1584
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
                    Left =2946
                    Top =1586
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
                    Left =9878
                    Top =1587
                    Width =585
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Pts"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Top =1133
                    Width =8226
                    Height =330
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="Field115"
                    ControlSource ="=[H_NAme] & ' - ' & [Sex Sub] & ' Results'"

                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =7553
                    Top =1587
                    Width =765
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
                    Left =8332
                    Top =1587
                    Width =1440
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text126"
                    Caption ="Result"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextFontFamily =18
                    Top =630
                    Width =3945
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text129"
                    Caption ="Competitor Results Summary"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =6298
                    Top =1587
                    Width =1020
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text130"
                    Caption ="Final"
                    FontName ="times New Roman"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            BreakLevel =1
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =30
            BreakLevel =3
            Name ="GroupHeader2"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10437
                    Name ="Line112"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =116
                    Width =2541
                    Height =285
                    FontSize =10
                    Name ="Fullname"
                    ControlSource ="Fullname"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =2777
                    Width =1011
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Age"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9921
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
                    Left =3797
                    Width =2376
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="ET_Des"
                    ControlSource ="ET_Des"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7491
                    Width =831
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="Field125"
                    ControlSource ="Place"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =8390
                    Width =1416
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="Field127"
                    ControlSource ="fResult"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =6236
                    Width =1086
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="Field131"
                    ControlSource ="F_Lev_Sub"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =90
                    Top =300
                    Width =10368
                    Name ="Line132"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =5616
                    Width =591
                    Height =285
                    FontSize =10
                    TabIndex =7
                    Name ="F_Lev"
                    ControlSource ="F_Lev"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =446
            OnFormat ="[Event Procedure]"
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
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter1"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons

Dim sHTML As String, rHTML As String, PageNum As Integer
Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
Dim ReportHead As String, GenerateHTML As Boolean

Const ReportTitle = "Competitor Results Ordered by House / Name"
Const repName = "coev" ' Keep to 4 letters or less (and unique from all other reports

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    Dim BGcolor As String

    If GenerateHTML And Not IsNull(Me!Age) And FormatCount = 1 Then
    
        DetailCount = DetailCount + 1
        
        If DetailCount = 1 Then

            If PageNum Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
            ' *** Create summary record ***
            Call RowStart(sHTML)

            Call CellStart(sHTML, "Center", "", "10%", BGcolor, 1)
            sHTML = sHTML & LinkStart(repName & PageNum & ".htm")
            Call Text(sHTML, "", "", str(PageNum))
            sHTML = sHTML & LinkEnd()
            Call CellEnd(sHTML)

            Call CellStart(sHTML, "Center", "", "30%", BGcolor, 1)
            Call Text(sHTML, "", "", Me![H_NAme])
            Call CellEnd(sHTML)
            
            Call CellStart(sHTML, "Center", "", "20%", BGcolor, 1)
            Call Text(sHTML, "", "", Me![Sex Sub])
            Call CellEnd(sHTML)
            
            Call CellStart(sHTML, "Left", "", "40%", BGcolor, 1)
            Call Text(sHTML, "", "", Me![Fullname] & " ...")
            Call CellEnd(sHTML)

            Call RowEnd(sHTML)

            ' *** Create general record header ***
            Call RowStart(rHTML)
            
            Call CellStart(rHTML, "Left", "Center", "25%", cCream, 1)
            Call SpaceIndent(rHTML, 2)
            Call Text(rHTML, "<B>", "</B>", "COMPETITOR")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "10%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "AGE")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "20%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "EVENT")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "10%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "FINAL")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "10%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "PLACE")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "10%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "RESULT")
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "Center", "5%", cCream, 1)
            Call Text(rHTML, "<B>", "</B>", "PTS")
            Call CellEnd(rHTML)
            
        End If

        'Debug.Print Me!Fullname, Me!et_des, Me!f_lev_sub, FormatCount

        
        If DetailCount Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
        If (Me!Place = 1) And (Me!F_Lev = 0) Then BGcolor = cLightRed

        Call RowStart(rHTML)
        
        Call CellStart(rHTML, "", "", "", BGcolor, 1)
        Call SpaceIndent(rHTML, 2)
        Call Text(rHTML, "", "", Me!Fullname)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Age)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!ET_Des)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!F_Lev_Sub)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Place)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!fResult)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "<B>", "</B>", Me!Points)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
    End If

End Sub

Private Sub PageFooter2_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    If GenerateHTML Then
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        
        Call TableEnd(rHTML)
        Call CreateHTMLfile(repName & PageNum & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & PageNum, ReportHead, repName)
    End If

End Sub

Private Sub PageHeader0_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    If GenerateHTML Then
        DetailCount = 0
        PageNum = PageNum + 1
        rHTML = ""
        
        If PageNum > 1 Then
            PrevPage = Link(repName & PageNum - 1 & ".htm", "Previous Page")
        Else
            PrevPage = ""
        End If
        NextPage = Link(repName & PageNum + 1 & ".htm", "Next Page")
        
        Call TableStart(rHTML, "95%", "", "", "", 0)
            
        ' *** Setup group title
        Call RowStart(rHTML)
        Call CellStart(rHTML, "Left", "Center", "100%", cWhite, 1)
        rHTML = rHTML & Heading(3, Me![H_NAme] & " - " & Me![Sex Sub] & " Results", 5)
        Call CellEnd(rHTML)
        Call RowEnd(rHTML)
        Call TableEnd(rHTML)

        Call TableStart(rHTML, "95%", "", "", "", 0)

    End If

End Sub

Private Sub Report_Close()
    On Error Resume Next

    If GenerateHTML Then
        GenerateHTML = False
        
        NextPage = ""
        Call TableEnd(rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
    
        Call CreateHTMLfile(repName & PageNum & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & PageNum, ReportHead, repName)
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead, repName)
    
        DoCmd.RunMacro "ClosePleaseWait"
        HTMLgenerateFinished = True
        
    End If
    
    DoCmd.RunMacro "ReportPopup-Update"
End Sub

Private Sub Report_Open(Cancel As Integer)
    
On Error Resume Next
    
    ' *** HTML Creation Code ***
    GenerateHTML = GlobalGenerateHTML
    
    If GenerateHTML Then
        PleaseWaitMsg = "Preparing HTML for """ & ReportTitle & """.  Please wait..."
        DoCmd.RunMacro "ShowPleaseWait"
        HTMLgenerateFinished = False
        
    End If
    
    PageNum = 0
    LastPage = False
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")
    ' ***************************

End Sub


Private Sub ReportHeader0_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code ***
    If GenerateHTML Then

        Call TableStart(sHTML, "90%", "", "", "", 0)
        Call RowStart(sHTML)

        Call CellStart(sHTML, "Center", "", "10%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "PAGE")
        Call CellEnd(sHTML)

        Call CellStart(sHTML, "Center", "", "30%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "TEAM")
        Call CellEnd(sHTML)
        
        Call CellStart(sHTML, "Center", "", "20%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "GENDER")
        Call CellEnd(sHTML)
        
        Call CellStart(sHTML, "Center", "", "40%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "COMPETITOR...")
        Call CellEnd(sHTML)

        Call RowEnd(sHTML)
    End If
    '*** ***************** ***

End Sub
