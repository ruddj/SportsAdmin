Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10656
    ItemSuffix =131
    Left =2970
    Top =1035
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xab9e2f379c64e240
    End
    RecordSource ="Statistics-Age Champion"
    Caption ="Statistics-Age Champion"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x350200003502000045020000d002000000000000a02900005401000001000000 ,
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
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="AgeSex"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SumOfPoints"
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
            Height =1305
            OnFormat ="[Event Procedure]"
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
                    Left =56
                    Top =566
                    Width =10596
                    Name ="Line74"
                End
                Begin Label
                    TextFontFamily =18
                    Left =113
                    Top =623
                    Width =5880
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Text122"
                    Caption ="Competitor Points - by Age / Gender"
                    FontName ="times New Roman"
                End
                Begin Label
                    Left =120
                    Top =1035
                    Width =9975
                    Height =255
                    Name ="Label130"
                    Caption ="The report tallies points that competitors gained for events in the age division"
                        " only. Points from other age divisions are not included."
                End
            End
        End
        Begin BreakHeader
            RepeatSection = NotDefault
            Height =925
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =1247
                    Top =565
                    Width =1365
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text123"
                    Caption ="Name"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =18
                    Left =4877
                    Top =557
                    Width =2130
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text124"
                    Caption ="Team"
                    FontName ="times New Roman"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =18
                    Left =8442
                    Top =503
                    Width =1455
                    Height =360
                    FontSize =13
                    FontWeight =700
                    Name ="Text125"
                    Caption ="Total"
                    FontName ="times New Roman"
                End
                Begin Line
                    BorderWidth =2
                    Left =506
                    Top =900
                    Width =10077
                    Name ="Line126"
                End
                Begin Label
                    TextFontFamily =18
                    Width =2205
                    Height =375
                    FontSize =14
                    FontWeight =700
                    Name ="Text127"
                    Caption ="AGE / GENDER:"
                    FontName ="times New Roman"
                End
                Begin TextBox
                    TextFontFamily =34
                    Left =2267
                    Top =50
                    Width =7986
                    Height =330
                    FontSize =11
                    FontWeight =700
                    Name ="AgeSex"
                    ControlSource ="AgeSex"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =355
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =1307
                    Width =3441
                    Height =330
                    FontSize =10
                    Name ="Fullname"
                    ControlSource ="Fullname"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4875
                    Width =2901
                    Height =330
                    FontSize =10
                    TabIndex =1
                    Name ="Field100"
                    ControlSource ="H_NAme"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =8496
                    Width =1521
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Field106"
                    ControlSource ="SumOfPoints"

                End
                Begin Line
                    Left =510
                    Top =340
                    Width =10077
                    Name ="Line86"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =510
                    Width =636
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Position"
                    ControlSource ="=1"

                End
            End
        End
        Begin BreakFooter
            Height =396
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =390
            OnFormat ="[Event Procedure]"
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
                    Width =1131
                    Height =390
                    FontSize =6
                    TabIndex =1
                    Name ="Field83"
                    ControlSource ="=DLookUp(\"[License]\",\"MiscellaneousLocal\")"

                End
                Begin TextBox
                    TextFontFamily =18
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

Dim HTM() As HTMarrayType

Const ReportTitle = "Age Champions"
Const repName = "agch" ' Keep to 4 letters or less (and unique from all other reports


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

On Error Resume Next

    If DisplayRecords >= NumberToDisplay Then
        Cancel = True
        Exit Sub
    End If

    DisplayRecords = DisplayRecords + 1


    '*** HTML Generation Code Start ***
    
    If GenerateHTML And Not Cancel And FormatCount = 1 Then
        
        DetailCount = DetailCount + 1
        
        If DetailCount Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
        
        rHTML = ""
        Call RowStart(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call SpaceIndent(rHTML, 2)
        Call Text(rHTML, "", "", Me!Position)
        Call CellEnd(rHTML)

        Call CellStart(rHTML, "", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!Fullname)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!H_NAme)
        Call CellEnd(rHTML)
        
        Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
        Call Text(rHTML, "", "", Me!SumOfPoints)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)

        
        Call AddToArray(Me!AgeSex, rDetail, rHTML)
    End If

    '*** HTML Generation Code End ***


End Sub

Private Sub Group1()

On Error Resume Next

    Call RowStart(rHTML)

    Call CellStart(rHTML, "", "", "10%", cWhite, 5)
    rHTML = rHTML & Heading(3, "Age / Gender: " & Me!AgeSex, 3)
    Call CellEnd(rHTML)
    
    Call RowEnd(rHTML)
    
    ' *** Create general record header ***
    Call RowStart(rHTML)
    
    Call CellStart(rHTML, "Center", "", "10%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "POS.")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "40%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "COMPETITOR")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "40%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TEAM")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "10%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TOTAL")
    Call CellEnd(rHTML)

    Call RowEnd(rHTML)

    Call AddToArray(Me!AgeSex, rGroupHeader, rHTML)

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    'DisplayRecords = 0

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        rHTML = ""
        Call RowStart(rHTML)
    
        Call CellStart(rHTML, "Center", "", "10%", cWhite, 1)
        Call CellEnd(rHTML)
        
        Call RowEnd(rHTML)
        
        Call AddToArray(Me!AgeSex, rGroupFooter, rHTML)

    End If

End Sub

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    DisplayRecords = 0

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        Call Group1
    End If

End Sub

Private Sub PageFooter2_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    '*** HTML Generation Code Start ***
    If GenerateHTML Then
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
                
        rHTML = ""
        Call TableEnd(rHTML)
        Call AddToArray(Me!AgeSex, rPageFooter, rHTML)
        
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
        
        Call AddToArray(Me!AgeSex, rPageHeader, rHTML)

    End If


End Sub


Private Sub Report_Close()

On Error Resume Next

Dim gHeader As Integer, OldPg As Integer, OldGroupName As String, i As Integer
Dim NewPg As Integer, CurrentGroupHeader As String
Dim eHTML As String, AlleHTML As String, sEvents As String

    If GenerateHTML Then
        

        GenerateHTML = False
        
        rHTML = ""
        Call TableEnd(rHTML)
    
        'Debug.Print "RF - FormatCount="; FormatCount; " Page="; PageNum;  me!Etdes
        Call AddToArray(Me!AgeSex, False, rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
        Call TableStart(sHTML, "90%", "", "", "", 0)
        Call RowStart(sHTML)

        Call CellStart(sHTML, "Center", "", "5%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "PAGE")
        Call CellEnd(sHTML)

        Call CellStart(sHTML, "Center", "", "80%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "AGE(s)")
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
            
            Debug.Print HTM(i).GrpHead; "|"; HTM(i).GrpName; "|"; HTM(i).Pg
            
            NewPg = HTM(i).Pg
            If HTM(i).GrpHead = rPageHeader Then
                If i = 2 Then Stop
                ' *** Create HTML Page
                rHTML = HTM(i).row & rHTML
                ' * Ensures that there is a header at the top of every page
                'If Not gHeader Then
                '    rHTML = CurrentGroupHeader & rHTML
                'End If
                
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
                Call CreateHTMLfile(repName & OldPg & ".htm", Template, rHTML, PrevPage, NextPage, ReportTitle & "  - Page " & OldPg, ReportHead, repName)
                rHTML = ""
                
                ' *** Create summary record ***
                If OldPg Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
                
                Call RowStart(eHTML)
    
                Call CellStart(eHTML, "Center", "", "20%", BGcolor, 1)
                eHTML = eHTML & LinkStart(repName & OldPg & ".htm")
                Call Text(eHTML, "", "", str(OldPg))
                eHTML = eHTML & LinkEnd()
                Call CellEnd(eHTML)
    
                Call CellStart(eHTML, "", "", "80%", BGcolor, 1)
                Call Text(eHTML, "", "", sEvents)
                Call CellEnd(eHTML)
                
                Call RowEnd(eHTML)
        
                AlleHTML = eHTML & AlleHTML
                eHTML = ""
                sEvents = ""
                
            End If
            
            If (HTM(i).GrpHead = rGroupHeader) And Not gHeader Then
                'If i = 2 Then Stop
                gHeader = True
                rHTML = HTM(i).row & rHTML
                
                Dim SpacedEvent As String

                SpacedEvent = HTM(i).GrpName
                Call SpaceIndent(SpacedEvent, 5)
                sEvents = SpacedEvent & " " & sEvents   ' * Adding each group title on page
                                                        ' *  to summary record
                'rHTML = HTM(i).row & rHTML              ' * Adding detail row

            End If
            
            ' *** Check for new group header ***
            If (OldGroupName <> HTM(i).GrpName) Then

                ' *** Add Group Header ***
                If (HTM(i).GrpHead <> rPageFooter) Then
                    gHeader = False
                
                End If
            End If

            ' *** Add Detail ***
            If OldGroupName = HTM(i).GrpName And Not gHeader Then
                rHTML = HTM(i).row & rHTML
            End If
            

            
            ' *** Set Old Group Header to current group header ***
            ' *** Ignore PageFooter groupType.  I hope it is not needed ever
            If (HTM(i).GrpHead <> rPageFooter) Then
                OldGroupName = HTM(i).GrpName
            End If
            OldPg = NewPg
        Next

        ' * Generate Summary Page file
        sHTML = sHTML & AlleHTML
        Call TableEnd(sHTML)
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead, repName)


    End If
    
    DoCmd.RunMacro "ReportPopup-Update"
End Sub

Private Sub Report_Open(Cancel As Integer)

On Error Resume Next

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

    DisplayRecords = 0
    NumberToDisplay = DLookup("[AgeChampionNumber]", "Misc-Statistics")
    If IsNull(NumberToDisplay) Then
      NumberToDisplay = 1
    End If
    Call UpdateEventCompetitorAge

End Sub
