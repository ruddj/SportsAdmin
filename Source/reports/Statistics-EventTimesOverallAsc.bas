Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridX =50
    GridY =50
    Width =10685
    ItemSuffix =142
    Left =30
    Top =120
    ShortcutMenuBar ="cmdReportRightClick"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0x18a2097f04e6e140
    End
    RecordSource ="Statistics-EventTimesOverallAsc"
    Caption ="Event Results - Best"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    PrtMip = Begin
        0x370200003702000045020000d002000000000000bd2900008c01000001000000 ,
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
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="Sex Sub"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            ControlSource ="OrderedResult"
        End
        Begin BreakLevel
            ControlSource ="Fullname"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader0"
        End
        Begin PageHeader
            Height =1133
            OnFormat ="[Event Procedure]"
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
                    ControlSource ="=\"Page \" & [Page]"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =866
            BreakLevel =2
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin Label
                    BackStyle =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =4988
                    Top =453
                    Width =2280
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text122"
                    Caption ="Event Description"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =1
                    TextFontFamily =18
                    Left =56
                    Top =451
                    Width =1365
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text107"
                    Caption ="Name"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =4144
                    Top =453
                    Width =630
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text108"
                    Caption ="Age"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =9751
                    Top =454
                    Width =825
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text111"
                    Caption ="Points"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =7937
                    Top =453
                    Width =675
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text124"
                    Caption ="Place"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =8736
                    Top =454
                    Width =1020
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text126"
                    Caption ="Result"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =3288
                    Top =454
                    Width =690
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text137"
                    Caption ="Team"
                    FontName ="Times New Roman"
                End
                Begin Label
                    BackStyle =0
                    TextAlign =2
                    TextFontFamily =18
                    Left =7370
                    Top =453
                    Width =570
                    Height =345
                    FontSize =11
                    FontWeight =700
                    Name ="Text138"
                    Caption ="Final"
                    FontName ="Times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Width =8226
                    Height =330
                    FontSize =12
                    FontWeight =700
                    Name ="ETdes"
                    ControlSource ="=[ET_Des] & ' - ' & [Sex Sub] & '  ' & [Age]"

                End
                Begin Line
                    BorderWidth =2
                    Left =29
                    Top =806
                    Width =10656
                    Name ="Line141"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =396
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7993
                    Top =14
                    Width =621
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="PlaceS"
                    ControlSource ="PlaceS"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =116
                    Top =14
                    Width =3111
                    Height =285
                    FontSize =10
                    Name ="Fullname"
                    ControlSource ="Fullname"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4155
                    Top =14
                    Width =831
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Age"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9993
                    Top =14
                    Width =531
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Points"
                    ControlSource ="Points"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =4977
                    Top =14
                    Width =2406
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="ET_Des"
                    ControlSource ="ET_Des"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =8618
                    Top =14
                    Width =1176
                    Height =285
                    FontSize =10
                    TabIndex =5
                    Name ="fResult"
                    ControlSource ="fResult"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =3240
                    Top =14
                    Width =966
                    Height =285
                    FontSize =10
                    TabIndex =6
                    Name ="H_Code"
                    ControlSource ="H_Code"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =7426
                    Top =14
                    Width =516
                    Height =285
                    FontSize =10
                    TabIndex =7
                    Name ="F_Lev"
                    ControlSource ="F_Lev"

                End
                Begin Line
                    OldBorderStyle =4
                    BorderLineStyle =3
                    Left =56
                    Top =340
                    Width =10482
                    Name ="Line112"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =2592
                    Width =276
                    Height =285
                    FontSize =10
                    TabIndex =8
                    Name ="Text140"
                    ControlSource ="PIN"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =226
            BreakLevel =2
            OnFormat ="[Event Procedure]"
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

Option Explicit

Dim DisplayRecords As Variant
Dim NumberToDisplay As Variant

' Generate HTML Variables and Constants
Dim sHTML As String, rHTML As String, PageNum As Integer, OldPg As Integer
Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
Dim ReportHead As String, GenerateHTML As Boolean, aIndex As Integer

Dim PreviousPIN As Variant

Dim HTM() As HTMarrayType

Const ReportTitle = "Event Results - Best"
Const repName = "etoa" ' Keep to 4 letters or less (and unique from all other reports


Private Sub AddToArray(GrpName As Variant, GrpHead As Integer, s As String)

On Error Resume Next

    aIndex = aIndex + 1
    
    ReDim Preserve HTM(aIndex) As HTMarrayType
    HTM(aIndex).Pg = PageNum
    HTM(aIndex).GrpName = GrpName
    HTM(aIndex).GrpHead = GrpHead
    HTM(aIndex).row = s

End Sub


Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)

  If PreviousPIN = Me!PIN Then
    Cancel = True
  Else
    On Error Resume Next

    If DisplayRecords >= NumberToDisplay Then
        Cancel = True
    End If

    DisplayRecords = DisplayRecords + 1


    '*** HTML Generation Code Start ***
    
    If GenerateHTML And Not Cancel And FormatCount = 1 Then
        
        DetailCount = DetailCount + 1
        
        If DetailCount Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
        
        rHTML = ""
        Call RowStart(rHTML)
        
        Call CellStart(rHTML, "Left", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Fullname)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        If Not IsNull(Me!H_Code) Then Call Text(rHTML, "", "", Me!H_Code)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Age)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!ET_Des)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!F_Lev)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!PlaceS)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!fResult)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Points)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        'Debug.Print "Detail - FormatCount="; FormatCount; " Page="; PageNum; me!Etdes
        Call AddToArray(Me!ETdes, rDetail, rHTML)
    End If

    '*** HTML Generation Code End ***
    
  End If
  
  PreviousPIN = Me!PIN

End Sub

Private Sub Group1()

On Error Resume Next

    '*** HTML Generation Code Start ***

    rHTML = ""
    ' *** Create Group Title
    Call RowStart(rHTML)

    Call CellStart(rHTML, "", "", "10%", cWhite, 5)
    rHTML = rHTML & Heading(3, Me!ETdes, 3)
    Call CellEnd(rHTML)
    
    Call RowEnd(rHTML)
    
    ' *** Create general record header ***
    Call RowStart(rHTML)
    
    Call CellStart(rHTML, "Center", "", "30%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "COMPETITOR")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "15%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TEAM")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "10%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "AGE")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "15%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "EVENT")
    Call CellEnd(rHTML)

    Call CellStart(rHTML, "Center", "", "5%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "FINAL")
    Call CellEnd(rHTML)

    Call CellStart(rHTML, "Center", "", "5%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "PLACE")
    Call CellEnd(rHTML)

    Call CellStart(rHTML, "Center", "", "15%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "RESULT")
    Call CellEnd(rHTML)

    Call CellStart(rHTML, "Center", "", "5%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "PTS")
    Call CellEnd(rHTML)

    Call RowEnd(rHTML)

    'Debug.Print "GH - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & me!Etdes
    Call AddToArray(Me!ETdes, rGroupHeader, rHTML)

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    DisplayRecords = 0

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        rHTML = ""
        Call RowStart(rHTML)
    
        Call CellStart(rHTML, "Center", "", "10%", cWhite, 1)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        Call AddToArray(Me!ETdes, rGroupFooter, rHTML)

    End If

End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next
    PreviousPIN = Null

    DisplayRecords = 0

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        Call Group1
    End If

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
  PreviousPIN = Null
End Sub

Private Sub PageFooter2_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
                
        rHTML = ""
        Call TableEnd(rHTML)
        Call AddToArray(Me!ETdes, rPageFooter, rHTML)
        
    End If


End Sub


Private Sub PageHeader0_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        'NewPage = True
        
        'DetailCount = 0
        PageNum = PageNum + 1
        rHTML = ""
        
        If PageNum > 1 Then
            PrevPage = Link(repName & PageNum - 1 & ".htm", "Previous Page")
        Else
            PrevPage = ""
        End If
        NextPage = Link(repName & PageNum + 1 & ".htm", "Next Page")
        
        Call TableStart(rHTML, "95%", "", "", "", 0)
        
        Call AddToArray(Me!ETdes, rPageHeader, rHTML)

    End If


End Sub

Private Sub Report_Close()

    On Error Resume Next
    
    Dim gHeader As Integer, OldPg As Integer, OldGroupName As String, i As Integer
    Dim NewPg As Integer

    If GenerateHTML Then
        Dim eHTML As String, AlleHTML As String, sEvents   As String

        GenerateHTML = False
        
        rHTML = ""
        Call TableEnd(rHTML)
    
        'Debug.Print "RF - FormatCount="; FormatCount; " Page="; PageNum;  me!Etdes
        Call AddToArray(Me!ETdes, False, rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
        Call TableStart(sHTML, "90%", "", "", "", 0)
        Call RowStart(sHTML)

        Call CellStart(sHTML, "Center", "", "5%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "PAGE")
        Call CellEnd(sHTML)

        Call CellStart(sHTML, "Center", "", "95%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "EVENT(S)")
        Call CellEnd(sHTML)
        
        Call RowEnd(sHTML)

    
        DoCmd.RunMacro "ClosePleaseWait"
        
        OldPg = HTM(aIndex).Pg
        gHeader = False
        OldGroupName = HTM(aIndex).GrpName
        
        ' Initialise variables to create Summary Page
        sEvents = OldGroupName
        eHTML = ""
        AlleHTML = ""
        
        rHTML = ""
        
        For i = aIndex To 1 Step -1
            
            NewPg = HTM(i).Pg
            If HTM(i).GrpHead = rPageHeader Then
                
                ' *** Create HTML Page
                rHTML = HTM(i).row & rHTML
                If OldPg > 1 Then
                    PrevPage = Link(repName & OldPg - 1 & ".htm", "Previous Page")
                Else
                    PrevPage = ""
                End If
                If OldPg < HTM(aIndex).Pg Then
                    NextPage = Link(repName & OldPg + 1 & ".htm", "Next Page")
                Else
                    NextPage = ""
                End If
                Call CreateHTMLfile(repName & OldPg & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & OldPg, ReportHead)
                rHTML = ""
                
                ' *** Create summary record ***
                If OldPg Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
                
                Call RowStart(eHTML)
    
                Call CellStart(eHTML, "Center", "", "5%", BGcolor, 1)
                eHTML = eHTML & LinkStart(repName & OldPg & ".htm")
                Call Text(eHTML, "", "", Str(OldPg))
                eHTML = eHTML & LinkEnd()
                Call CellEnd(eHTML)
    
                Call CellStart(eHTML, "", "", "95%", BGcolor, 1)
                Call Text(eHTML, "", "", sEvents)
                Call CellEnd(eHTML)
                
                Call RowEnd(eHTML)
        
                AlleHTML = eHTML & AlleHTML
                eHTML = ""
                sEvents = ""

            End If
            
            If (HTM(i).GrpHead = rGroupHeader) And Not gHeader Then
                gHeader = True
                rHTML = HTM(i).row & rHTML
            End If
            
            If OldGroupName = HTM(i).GrpName And Not gHeader Then
                rHTML = HTM(i).row & rHTML
            
            ElseIf (OldGroupName <> HTM(i).GrpName) And (HTM(i).GrpHead <> rPageFooter) Then
                Dim SpacedEvent As String

                SpacedEvent = HTM(i).GrpName
                Call SpaceIndent(SpacedEvent, 5)
                sEvents = SpacedEvent & " " & sEvents
                rHTML = HTM(i).row & rHTML
                gHeader = False
            End If
            
            If (HTM(i).GrpHead = rGroupHeader) And (OldGroupName <> HTM(i).GrpName) Then
                gHeader = True
                rHTML = HTM(i).row & rHTML
            End If
            
            
            'Debug.Print HTM(i).Pg, HTM(i).GrpName, HTM(i).GrpHead', HTM(i).row

            ' Ignore PageFooter groupType.  I hope it is not needed ever
            If (HTM(i).GrpHead <> rPageFooter) Then
                OldGroupName = HTM(i).GrpName
            End If
            OldPg = NewPg
        Next

        ' * Generate Summary Page file
        sHTML = sHTML & AlleHTML
        Call TableEnd(sHTML)
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead)


    End If
    
    DoCmd.RunMacro "ReportPopup-Update"
    
End Sub

Private Sub Report_Open(Cancel As Integer)

PreviousPIN = Null

On Error Resume Next

    DisplayRecords = 0
    'NumberToDisplay = DLookup("[CompetitorPlaces]", "MiscellaneousLocal")
    NumberToDisplay = DLookup("[NumberOfRecords]", "Misc-Statistics")
    ' *** HTML Creation Code ***
    GenerateHTML = GlobalGenerateHTML
    
    If GenerateHTML Then
        aIndex = 0
        PleaseWaitMsg = "Preparing HTML for """ & ReportTitle & """.  Please wait..."
        DoCmd.RunMacro "ShowPleaseWait"
    End If
    
    PageNum = 0
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")
    ' ***************************

End Sub
