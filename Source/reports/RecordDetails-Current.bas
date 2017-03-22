Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =10714
    ItemSuffix =142
    Left =1530
    Top =15
    OnNoData ="[Event Procedure]"
    ShortcutMenuBar ="Sports Admin-Print Popup"
    Toolbar ="Sports Admin-Print"
    RecSrcDt = Begin
        0xeb4e8a91cedae140
    End
    RecordSource ="Report-Records"
    OnOpen ="[Event Procedure]"
    OnClose ="ReportPopup-Update"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x370200003702000045020000d002000000000000da2900001801000001000000 ,
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
            KeepTogether =1
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Sex"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="BestResult"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader0"
        End
        Begin PageHeader
            Height =971
            OnFormat ="[Event Procedure]"
            Name ="PageHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Width =10656
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
                    Width =10596
                    Name ="Line74"
                End
                Begin Label
                    TextFontFamily =18
                    Left =56
                    Top =566
                    Width =5880
                    Height =405
                    FontSize =15
                    FontWeight =700
                    Name ="Text102"
                    Caption ="RECORD HOLDERS"
                    FontName ="times New Roman"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =738
            OnFormat ="[Event Procedure]"
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Left =56
                    Top =56
                    Width =8346
                    Height =285
                    FontSize =11
                    FontWeight =700
                    Name ="Event"
                    ControlSource ="=\"EVENT: \" & [ET_Des]"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextFontFamily =34
                    Left =283
                    Top =453
                    Width =960
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text126"
                    Caption ="DIVISION"
                End
                Begin Label
                    TextFontFamily =34
                    Left =2211
                    Top =453
                    Width =660
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text127"
                    Caption ="NAME"
                End
                Begin Label
                    TextFontFamily =34
                    Left =5612
                    Top =453
                    Width =645
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text128"
                    Caption ="TEAM"
                End
                Begin Label
                    TextFontFamily =34
                    Left =7483
                    Top =453
                    Width =885
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text129"
                    Caption ="RESULT"
                End
                Begin Label
                    TextFontFamily =34
                    Left =9411
                    Top =453
                    Width =615
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text130"
                    Caption ="DATE"
                End
                Begin Line
                    BorderWidth =2
                    Left =56
                    Top =396
                    Width =10545
                    Name ="Line140"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =9216
                    Width =876
                    Height =225
                    FontSize =9
                    TabIndex =1
                    Name ="ETdes"
                    ControlSource ="ET_Des"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader1"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =2
            OnFormat ="[Event Procedure]"
            Name ="AgeHeader"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =280
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =2199
                    Width =3231
                    Height =225
                    FontSize =9
                    Name ="FullName"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =5604
                    Width =1716
                    Height =225
                    FontSize =9
                    TabIndex =1
                    Name ="H_NAme"
                    ControlSource ="CompetitorHouse"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =285
                    Width =801
                    Height =225
                    FontSize =9
                    TabIndex =2
                    Name ="Age"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =7486
                    Width =846
                    Height =225
                    FontSize =9
                    TabIndex =3
                    Name ="Record"
                    ControlSource ="Result"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =8400
                    Width =906
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="Units"
                    ControlSource ="Units"

                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    Left =1140
                    Width =906
                    Height =225
                    FontSize =9
                    TabIndex =5
                    Name ="Sex Sub"
                    ControlSource ="Sex Sub"
                    EventProcPrefix ="Sex_Sub"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =9424
                    Width =846
                    Height =255
                    FontSize =9
                    TabIndex =6
                    Name ="Date"
                    ControlSource ="Date"

                End
                Begin Rectangle
                    BackStyle =0
                    Left =165
                    Width =1927
                    Height =280
                    Name ="Box135"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =2092
                    Width =3397
                    Height =280
                    Name ="Box136"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =5484
                    Width =1882
                    Height =280
                    Name ="Box137"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =7370
                    Width =1927
                    Height =280
                    Name ="Box138"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =9302
                    Width =1072
                    Height =280
                    Name ="Box139"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            BreakLevel =2
            OnFormat ="[Event Procedure]"
            Name ="AgeFooter"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =113
            BreakLevel =1
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter2"
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
        End
        Begin PageFooter
            Height =446
            OnFormat ="[Event Procedure]"
            Name ="PageFooter2"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =9467
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
                    Top =56
                    Width =10596
                    Name ="Line87"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            OnFormat ="[Event Procedure]"
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

' Generate HTML Variables and Constants
Dim sHTML As String, rHTML As String, PageNum As Integer, OldPg As Integer
Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
Dim ReportHead As String, GenerateHTML As Boolean, aIndex As Integer
Dim AgeCount As Integer, PreviousResult As Variant

Dim HTM() As HTMarrayType

Const ReportTitle = "Record Holders"
Const repName = "rh" ' Keep to 4 letters or less (and unique from all other reports

Private Sub AddToArray(GrpName As Variant, GrpHead As Integer, s As String)

On Error Resume Next

    aIndex = aIndex + 1
    
    ReDim Preserve HTM(aIndex) As HTMarrayType
    HTM(aIndex).Pg = PageNum
    HTM(aIndex).GrpName = GrpName
    HTM(aIndex).GrpHead = GrpHead
    HTM(aIndex).row = s

End Sub

Private Sub AgeFooter_Format(Cancel As Integer, FormatCount As Integer)
  
  AgeCount = 1
  
End Sub

Private Sub AgeHeader_Format(Cancel As Integer, FormatCount As Integer)
  
  AgeCount = 0
  
End Sub

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)

  AgeCount = AgeCount + 1
  If (AgeCount > 1) And (PreviousResult <> Me!Result) Then
    Cancel = True
  Else
    PreviousResult = Me!Result
    On Error Resume Next
  
      '*** HTML Generation Code Start ***
      
      If GenerateHTML And Not Cancel And FormatCount = 1 Then
          
          DetailCount = DetailCount + 1
          
          If DetailCount Mod 2 = 0 Then BGcolor = cWhite Else BGcolor = cLightGray
          
          rHTML = ""
          Call RowStart(rHTML)
          
          Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
          Call SpaceIndent(rHTML, 2)
          Call Text(rHTML, "", "", Me!Age)
          Call CellEnd(rHTML)
          
          Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
          Call Text(rHTML, "", "", Me![Sex Sub])
          Call CellEnd(rHTML)
          
          Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
          Call Text(rHTML, "", "", Me!Fullname)
          Call CellEnd(rHTML)
          
          Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
          Call Text(rHTML, "", "", Me!H_NAme)
          Call CellEnd(rHTML)
          
          Call CellStart(rHTML, "Right", "", "", BGcolor, 1)
          Call Text(rHTML, "", "", Me!Record)
          Call CellEnd(rHTML)
          
          Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
          Call Text(rHTML, "", "", Me!Units)
          Call CellEnd(rHTML)
          
          Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
          Call Text(rHTML, "", "", Me!Date)
          Call CellEnd(rHTML)
          
          Call RowEnd(rHTML)
          
          'Debug.Print "Detail - FormatCount="; FormatCount; " Page="; PageNum; Me!ET_Des
          Call AddToArray(Me!ETdes, rDetail, rHTML)
      End If
  
      '*** HTML Generation Code End ***
  End If
'Detail1_Format_Exit:
'    Exit Sub

'Detail1_Format_Err:
'    Resume Next
    
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
    
    Call CellStart(rHTML, "Center", "", "20%", cCream, 2)
    Call Text(rHTML, "<B>", "</B>", "DIVISION")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "35%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "COMPETITOR")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "20%", cCream, 0)
    Call Text(rHTML, "<B>", "</B>", "TEAM")
    Call CellEnd(rHTML)
    
    Call CellStart(rHTML, "Center", "", "15%", cCream, 2)
    Call Text(rHTML, "<B>", "</B>", "RESULT")
    Call CellEnd(rHTML)

    Call CellStart(rHTML, "Center", "", "10%", cCream, 2)
    Call Text(rHTML, "<B>", "</B>", "DATE")
    Call CellEnd(rHTML)

    Call RowEnd(rHTML)

    'Debug.Print "GH - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
    Call AddToArray(Me!ETdes, rGroupHeader, rHTML)

End Sub

Private Sub GroupFooter0_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

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

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)
    
On Error Resume Next

    'Debug.Print "GF - FormatCount="; FormatCount; " Page="; PageNum; " Event: " & Me!ET_Des
    
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

Private Sub Report_NoData(Cancel As Integer)

    MsgBox ("There is no data to display for the report: " & Me.Caption)
  Cancel = True

End Sub

Private Sub Report_Open(Cancel As Integer)

  AgeCount = 1
  
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


End Sub

Private Sub ReportFooter1_Format(Cancel As Integer, FormatCount As Integer)

On Error Resume Next

    If GenerateHTML Then
        Dim eHTML As String, AlleHTML As String, sEvents   As String

        GenerateHTML = False
        
        rHTML = ""
        Call TableEnd(rHTML)
    
        'Debug.Print "RF - FormatCount="; FormatCount; " Page="; PageNum;  Me!ETdes
        Call AddToArray(Me!ETdes, False, rHTML)
        
        Template = DLookup("[TemplateFile]", "MiscHTML")
        TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
        Call TableStart(sHTML, "90%", "", "", "", 0)
        Call RowStart(sHTML)

        Call CellStart(sHTML, "Center", "", "10%", cCream, 1)
        Call Text(sHTML, "<B>", "</B>", "PAGE")
        Call CellEnd(sHTML)

        Call CellStart(sHTML, "Center", "", "90%", cCream, 1)
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
    
                Call CellStart(eHTML, "Center", "", "20%", BGcolor, 1)
                eHTML = eHTML & LinkStart(repName & OldPg & ".htm")
                Call Text(eHTML, "", "", Str(OldPg))
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

End Sub
