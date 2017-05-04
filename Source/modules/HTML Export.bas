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


Private Sub TestExportNamesHTML()
    Dim sReport As String
    
    sReport = "rh"
    
    'sReport = "agca"
    Call ExportNamesHTML(sReport)


End Sub

Public Function ExportNamesHTML(Optional repName As String = "agca")
    ' Exports a list of competitors names in a formatted HTML page
    ' The HTML is generated based on the options defined in tblReportsHTML
    
    ' Version 2 of report. Try to use modern CSS output and more logical coding.
    ' Also try to abstract out report details to allow it to be used for multiple reports
    
    'On Error Resume Next
    
    Dim MyDb As Database, Rs As Recordset, QryName As String
    Dim curGroup As String, iPosition As Integer, iDisplayMax As Integer
    Dim bAgeChamp As Boolean, iDisplayLimit As Integer
    Dim ReportTitle As String, ReportCaption As Variant, repGroup As String, repGroupHeader As String
    Dim repFinalLev As Variant, repPlace As Variant, strPlace As String
       
    Dim dataFields() As String, dataHeaders() As String
    Dim varField As Variant, varValue As Variant, strField As String, strValue As String
    Dim cssGroup As String
    
    Dim sHTML As String ' Summary and Shortcuts
    Dim pHTML As String ' Page Header
    Dim rHTML As String ' Results
    
    
    ' Check query definition exists
    If (IsNull(DLookup("[repQuery]", "tblReportsHTML", "[repShortCode] = """ & repName & """"))) Then
        MsgBox "No matching query entry found"
        Exit Function
    End If
    
    
    ' Code maybe parametised
    QryName = DLookup("[repQuery]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    ReportTitle = DLookup("[repTitle]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    ReportCaption = DLookup("[repCaption]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    repGroup = DLookup("[repGroup]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    If (IsNull(DLookup("[repGroupHeader]", "tblReportsHTML", "[repShortCode] = """ & repName & """"))) Then
        repGroupHeader = ""
    Else
        repGroupHeader = DLookup("[repGroupHeader]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    End If
    
    ' Load Fields
    dataFields = Split(DLookup("[repFields]", "tblReportsHTML", "[repShortCode] = """ & repName & """"), ";")
    dataHeaders = Split(DLookup("[repHeaders]", "tblReportsHTML", "[repShortCode] = """ & repName & """"), ";")
    
    
    If (UBound(dataFields) <> UBound(dataHeaders)) Then
        MsgBox "Report Fields and Headers do not match"
        Exit Function
    End If
    
    ' Is this an Age Championship
    iDisplayLimit = DLookup("[repDisplayLimit]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    bAgeChamp = DLookup("[repAgeChamp]", "tblReportsHTML", "[repShortCode] = """ & repName & """")

    If (iDisplayLimit > 0) Then
        'Display Limit is hard coded
        iDisplayMax = iDisplayLimit
    Else
        'Display Limit is set by user
        
        If (bAgeChamp) Then
            iDisplayMax = DLookup("[AgeChampionNumber]", "Misc-Statistics")
        Else
            iDisplayMax = DLookup("[NumberOfRecords]", "Misc-Statistics")
        End If
    End If
    

    repFinalLev = DLookup("[repFinalLev]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
    repPlace = DLookup("[repPlace]", "tblReportsHTML", "[repShortCode] = """ & repName & """")
         
    ' Load Data
    Set MyDb = CurrentDb()
    Set Rs = MyDb.OpenRecordset(QryName, dbOpenDynaset)

    ' Start Web Page Header
    ReportHead = DLookup("[ReportHeader]", "MiscHTML")
        
    ' Name report based on
    Call DivOpen(pHTML, "", "html-" & repName)
    
    Call DivOpen(pHTML, "header")
    pHTML = pHTML & Heading(1, ReportTitle, 0)
    
    If (Not IsNull(ReportCaption) And ReportCaption <> "") Then
        Call DivOpen(pHTML, "caption")
        pHTML = pHTML & Heading(3, ReportCaption, 0)
        Call DivClose(pHTML)
    End If
    
    Call DivClose(pHTML)
    
    ' Start Summary
    Call DivOpen(sHTML, "", "summary")
    sHTML = sHTML & Heading(2, "Summary of Results", 0)
    Call ListOpen(sHTML, "main")
    
    ' Start Results
    Call DivOpen(rHTML, "results")
       
    If (Rs.EOF Or Rs.BOF) Then
        ' No Data
        MsgBox "No Records for HTML Export"
        Exit Function
    End If
    
    ' Cycle through Data and add to array
    Rs.MoveFirst 'Unnecessary in this case, but still a good habit
    Do Until Rs.EOF = True
        ' If start of new group add entry
        If curGroup <> Rs(repGroup) Then
            If Rs.AbsolutePosition > 0 Then
                ' Not First Group add end group
                Call TableEnd(rHTML)
                Call DivClose(rHTML) ' Close Data
                Call DivOpen(rHTML, "grp-return")
                rHTML = rHTML & Link("#summary", "Return to Summary")
                Call DivClose(rHTML)
                Call DivClose(rHTML) ' Close Group
            End If
            curGroup = Rs(repGroup)
            iPosition = 1
            
            cssGroup = AlphaNumericDashOnly(curGroup)
            
            Call DivOpen(rHTML, "grp-results", "grp-" & cssGroup)
            Call ListItem(sHTML, Link("#grp-" & cssGroup, curGroup))
            'sHTML = sHTML & "" & Link("#grp-" & cssGroup, curGroup)
            Call DivOpen(rHTML, "hdr-" & cssGroup)
            rHTML = rHTML & Heading(3, repGroupHeader & " " & curGroup, 0)
            Call DivClose(rHTML)
            
            ' *** Create general record header ***
            Call DivOpen(rHTML, "data-" & cssGroup)
            Call TableOpen(rHTML, cssGroup)

            Call TableHeadOpen(rHTML, "")
            Call RowStart(rHTML)
            'groupHeader
            For Each varField In dataHeaders
                strField = CStr(varField)
                Call CellHead(rHTML, strField, StrConv(strField, vbLowerCase))
            Next varField
            
            Call RowEnd(rHTML)
            Call TableHeadEnd(rHTML)
        
        End If
        
        
        ' Add individual student to export
        If iPosition <= iDisplayMax Then
            strPlace = ""
            If (Not bAgeChamp And Not IsNull(repFinalLev) And Not IsNull(repPlace)) Then
                strPlace = "place-" & Trim(Rs(repFinalLev)) & "-" & Trim(Rs(repPlace))
            Else
                strPlace = "position-" & iPosition
            End If

            Call RowStart(rHTML, strPlace)
            
            For Each varField In dataFields
                strField = CStr(varField)
                If (strField = "_Position") Then
                    strValue = iPosition
                Else
                    varValue = Rs(varField)
                    If IsNull(varValue) Then
                        strValue = ""
                    Else
                        strValue = CStr(varValue)
                    End If
                End If
                Call Cell(rHTML, strValue, strField)
            Next varField
                        
            Call RowEnd(rHTML)
    
        End If
        'Move to the next record. Don't ever forget to do this.
        iPosition = iPosition + 1
        Rs.MoveNext
    Loop
    
    ' Close Last group
    Call TableEnd(rHTML)
    Call DivClose(rHTML) ' Close Data
    Call DivOpen(rHTML, "grp-return")
    rHTML = rHTML & Link("#summary", "Return to Summary")
    Call DivClose(rHTML)
    Call DivClose(rHTML) ' Close Group
    Call DivClose(rHTML) ' Close Results
    Call DivClose(rHTML) ' Close Report
    
    Debug.Print "End Loop " & repName
    
    ' Output HTML
    Template = DLookup("[TemplateFile]", "MiscHTML")
    TemplateSummary = DLookup("[TemplateFileSummary]", "MiscHTML")
    
    Call ListClose(sHTML)
    Call DivClose(sHTML)

    Call CreateHTMLfile(repName & ".htm", Template, pHTML & sHTML & rHTML, "", "", ReportTitle, ReportHead, repName)


End Function