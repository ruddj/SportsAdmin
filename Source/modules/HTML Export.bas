Option Compare Database

Option Explicit

Dim PageNum As Integer, OldPg As Integer
Dim ReportHead As String, aIndex As Integer
Dim HTM() As HTMarrayType

Private Sub AddToArray(GrpName As Variant, GrpHead As Integer, s As String)

On Error Resume Next

    aIndex = aIndex + 1
    
    ReDim Preserve HTM(aIndex) As HTMarrayType
    HTM(aIndex).Pg = PageNum
    HTM(aIndex).GrpName = GrpName
    HTM(aIndex).GrpHead = GrpHead
    HTM(aIndex).row = s

End Sub

Function AgeChampionAll()
    Dim MyDb As Database, rs As Recordset, QryName As String
    Dim curGroup As String, iPosition As Integer, iDisplayMax As Integer
    Dim sHTML As String, rHTML As String
    Dim gHeader As Integer, OldPg As Integer, OldGroupName As String, i As Integer
    Dim NewPg As Integer, CurrentGroupHeader As String
    Dim eHTML As String, AlleHTML As String, sEvents As String
    Dim LastPage As Integer, DetailCount As Integer, NextPage As String, PrevPage As String
        
    rHTML = ""
    curGroup = ""
    PageNum = 1
    
    ' Code maybe parametised
    QryName = "Statistics-Age Champion-AnyDivision"
    Const ReportTitle = "Age Champions"
    Const repName = "agca"
    iDisplayMax = DLookup("[AgeChampionNumber]", "Misc-Statistics")
      
    ' Load Data
    Set MyDb = CurrentDb()
    Set rs = MyDb.OpenRecordset(QryName, dbOpenDynaset)

    ' Start Web Page Header
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")
    Call TableStart(rHTML, "95%", "", "", "", 0)
    Call AddToArray(rs!AgeSex, rPageHeader, rHTML)
    
    
    If (rs.EOF And rs.BOF) Then
        ' No Data
        MsgBox "No Records for HTML Export"
        Exit Function
    End If
    ' Cycle through Data and add to array
    rs.MoveFirst 'Unnecessary in this case, but still a good habit
    Do Until rs.EOF = True
        ' If start of new group add entry
        If curGroup <> rs!AgeSex Then
            If rs.AbsolutePosition > 0 Then
                ' Not First Group add end group
                rHTML = ""
                Call RowStart(rHTML)
            
                Call CellStart(rHTML, "Center", "", "10%", cWhite, 1)
                Call CellEnd(rHTML)
                
                Call RowEnd(rHTML)
                
                Call AddToArray(curGroup, rGroupFooter, rHTML)
            
            End If
            curGroup = rs!AgeSex
            iPosition = 1
            rHTML = ""
        
            Call RowStart(rHTML)
        
            Call CellStart(rHTML, "", "", "10%", cWhite, 5)
            rHTML = rHTML & Heading(3, "Age / Gender: " & rs!AgeSex, 3)
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
        
            Call AddToArray(rs!AgeSex, rGroupHeader, rHTML)
        End If
        
        
        ' Add individual student to export
        If iPosition <= iDisplayMax Then
            rHTML = ""
            Call RowStart(rHTML)
            
            Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
            Call SpaceIndent(rHTML, 2)
            Call Text(rHTML, "", "", iPosition)
            Call CellEnd(rHTML)
    
            Call CellStart(rHTML, "", "", "", BGcolor, 1)
            Call Text(rHTML, "", "", rs!Fullname)
            Call CellEnd(rHTML)
            
            Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
            Call Text(rHTML, "", "", rs!H_NAme)
            Call CellEnd(rHTML)
    
            Call CellStart(rHTML, "Center", "", "", BGcolor, 1)
            Call Text(rHTML, "", "", rs!SumOfPoints)
            Call CellEnd(rHTML)
            
            Call RowEnd(rHTML)
    
            Call AddToArray(rs!AgeSex, rDetail, rHTML)
        End If
        'Move to the next record. Don't ever forget to do this.
        iPosition = iPosition + 1
        rs.MoveNext
    Loop
    
    ' Close Last group
    rHTML = ""
    Call RowStart(rHTML)

    Call CellStart(rHTML, "Center", "", "10%", cWhite, 1)
    Call CellEnd(rHTML)
    Call RowEnd(rHTML)
    
    Call AddToArray(curGroup, rGroupFooter, rHTML)
    
    Debug.Print "End Loop"
    rHTML = ""
    Call TableEnd(rHTML)
    Call AddToArray(curGroup, False, rHTML)
    
    ' Output HTML
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
        Call CreateHTMLfile("_" & repName & ".htm", TemplateSummary, sHTML, PrevPage, NextPage, "Summary of " & ReportTitle, ReportHead)



End Function