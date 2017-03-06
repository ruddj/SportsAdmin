Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =141
    Left =2190
    Top =300
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x5e4c82792b2ce240
    End
    RecordSource ="SELECT DISTINCTROW House.H_NAme, [Surname] & \", \" & [Gname] AS Fullname, Compe"
        "titors.PIN, [Sex Sub].[Sex Sub], Events.Age, EventType.ET_Des, CompEvents.Points"
        ", [Result] & ' ' & [Units] AS fResult, CompEvents.Place, [Final Level Sub].F_Lev"
        "_Sub, CompEvents.F_Lev FROM House RIGHT JOIN (EventType RIGHT JOIN ((Competitors"
        " LEFT JOIN [Sex Sub] ON Competitors.Sex = [Sex Sub].Sex) LEFT JOIN ((Events RIGH"
        "T JOIN CompEvents ON Events.E_Code = CompEvents.E_Code) LEFT JOIN [Final Level S"
        "ub] ON CompEvents.F_Lev = [Final Level Sub].F_Lev) ON Competitors.PIN = CompEven"
        "ts.PIN) ON EventType.ET_Code = Events.ET_Code) ON House.H_Code = Competitors.H_C"
        "ode WHERE (((EventType.Include)=True) AND ((EventType.Flag)=True) AND ((Events.I"
        "nclude)=True) AND ((House.Include)=True) AND ((House.Flag)=True)) ORDER BY [Surn"
        "ame] & \", \" & [Gname], CompEvents.Points DESC;"
    OnOpen ="[Event Procedure]"
    OnClose ="ReportPopup-Update"
    PrtMip = Begin
        0x3702000037020000450200009602000000000000a0290000fe01000001000000 ,
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
        Begin PageHeader
            Height =1474
            OnFormat ="[Event Procedure]"
            Name ="PageHeader0"
            Begin
                Begin Label
                    TextFontFamily =18
                    Top =105
                    Width =3945
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text129"
                    Caption ="Competitor Results"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =113
                    Top =680
                    Width =5331
                    Height =330
                    FontSize =12
                    Name ="Fullname"
                    ControlSource ="Fullname"

                End
                Begin Label
                    Left =120
                    Top =1245
                    Width =465
                    Height =225
                    Name ="Label135"
                    Caption ="Event"
                End
                Begin Label
                    Left =4485
                    Top =1215
                    Width =465
                    Height =225
                    Name ="Label136"
                    Caption ="Age"
                End
                Begin Label
                    Left =5610
                    Top =1185
                    Width =810
                    Height =225
                    Name ="Label137"
                    Caption ="Final Level"
                End
                Begin Label
                    Left =7425
                    Top =1170
                    Width =810
                    Height =225
                    Name ="Label138"
                    Caption ="Palce"
                End
                Begin Label
                    Left =8385
                    Top =1170
                    Width =810
                    Height =225
                    Name ="Label139"
                    Caption ="Resut"
                End
                Begin Label
                    Left =9690
                    Top =1170
                    Width =810
                    Height =225
                    Name ="Label140"
                    Caption ="Points"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =30
            BreakLevel =1
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
            Height =510
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4478
                    Top =90
                    Width =1011
                    Height =285
                    FontSize =10
                    Name ="Age"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9921
                    Top =90
                    Width =531
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Field106"
                    ControlSource ="Points"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =113
                    Top =90
                    Width =3861
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="ET_Des"
                    ControlSource ="ET_Des"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7491
                    Top =90
                    Width =831
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Field125"
                    ControlSource ="Place"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =8390
                    Top =90
                    Width =1416
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="Field127"
                    ControlSource ="fResult"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =5606
                    Top =90
                    Width =1716
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="Field131"
                    ControlSource ="F_Lev_Sub"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =90
                    Top =450
                    Width =10368
                    Name ="Line132"
                End
            End
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
            Call Text(sHTML, "", "", Str(PageNum))
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
        Call CreateHTMLfile(repName & PageNum & ".htm", Template, rHTML, PrevPage, NextPage, Heading(3, ReportTitle & "  - Page " & PageNum, 0), ReportHead)
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

Private Sub ReportFooter1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    If GenerateHTML Then
        GenerateHTML = False
        
        NextPage = ""
        Call TableEnd(rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
    
        Call CreateHTMLfile(repName & PageNum & ".htm", Template, rHTML, PrevPage, NextPage, Heading(3, ReportTitle & "  - Page " & PageNum, 0), ReportHead)
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead)
    
        DoCmd.RunMacro "ClosePleaseWait"
        HTMLgenerateFinished = True
        
    End If

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
