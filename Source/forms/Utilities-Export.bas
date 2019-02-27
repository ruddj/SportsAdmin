Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =5546
    ItemSuffix =71
    Left =-20400
    Top =4320
    Right =-11385
    Bottom =8955
    HelpContextId =565
    RecSrcDt = Begin
        0x6bd443042dc7e140
    End
    RecordSource ="Miscellaneous"
    Caption ="Utilities"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin OptionGroup
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin ToggleButton
            TextFontFamily =2
            Width =283
            Height =283
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =4251
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =113
                    Top =163
                    Width =5336
                    Height =2839
                    Name ="Box50"
                    LayoutCachedLeft =113
                    LayoutCachedTop =163
                    LayoutCachedWidth =5449
                    LayoutCachedHeight =3002
                End
                Begin Label
                    OverlapFlags =215
                    Left =226
                    Top =226
                    Width =2092
                    Height =220
                    FontSize =9
                    FontWeight =700
                    Name ="Label56"
                    Caption ="Meet Manager Export"
                    LayoutCachedLeft =226
                    LayoutCachedTop =226
                    LayoutCachedWidth =2318
                    LayoutCachedHeight =446
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1247
                    Top =570
                    Width =3126
                    Name ="Mteam"
                    ControlSource ="Mteam"

                    LayoutCachedLeft =1247
                    LayoutCachedTop =570
                    LayoutCachedWidth =4373
                    LayoutCachedHeight =810
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =285
                            Top =574
                            Width =975
                            Height =240
                            Name ="Label57"
                            Caption ="Team Name:"
                            LayoutCachedLeft =285
                            LayoutCachedTop =574
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =814
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1247
                    Top =967
                    Width =606
                    TabIndex =1
                    Name ="Mcode"
                    ControlSource ="Mcode"

                    LayoutCachedLeft =1247
                    LayoutCachedTop =967
                    LayoutCachedWidth =1853
                    LayoutCachedHeight =1207
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =285
                            Top =964
                            Width =945
                            Height =240
                            Name ="Label58"
                            Caption ="Team Code:"
                            LayoutCachedLeft =285
                            LayoutCachedTop =964
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1204
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =3911
                    Top =963
                    Width =516
                    TabIndex =2
                    Name ="Mtop"
                    ControlSource ="Mtop"

                    LayoutCachedLeft =3911
                    LayoutCachedTop =963
                    LayoutCachedWidth =4427
                    LayoutCachedHeight =1203
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =1980
                            Top =960
                            Width =1995
                            Height =240
                            Name ="Label59"
                            Caption ="# Top Competitors/Event:"
                            LayoutCachedLeft =1980
                            LayoutCachedTop =960
                            LayoutCachedWidth =3975
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontFamily =34
                    Left =1474
                    Top =2267
                    Width =1029
                    Height =570
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    Name ="MMExportEntry"
                    Caption ="Semi-Colon\015\012Delimited"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1474
                    LayoutCachedTop =2267
                    LayoutCachedWidth =2503
                    LayoutCachedHeight =2837
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontFamily =34
                    Left =2761
                    Top =2260
                    Width =1134
                    Height =600
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    Name ="MMExportComp"
                    Caption ="Semi-Colon\015\012Delimited"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =2761
                    LayoutCachedTop =2260
                    LayoutCachedWidth =3895
                    LayoutCachedHeight =2860
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =247
                    Left =226
                    Top =510
                    Width =4301
                    Height =799
                    Name ="Box61"
                    LayoutCachedLeft =226
                    LayoutCachedTop =510
                    LayoutCachedWidth =4527
                    LayoutCachedHeight =1309
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =170
                    Top =3628
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    HelpContextId =410
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =170
                    LayoutCachedTop =3628
                    LayoutCachedWidth =1304
                    LayoutCachedHeight =4138
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =223
                    TextFontFamily =34
                    Left =4021
                    Top =2260
                    Width =1194
                    Height =600
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    Name ="MMExportRE1"
                    Caption ="RE1 Registrations"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =4021
                    LayoutCachedTop =2260
                    LayoutCachedWidth =5215
                    LayoutCachedHeight =2860
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =170
                    Top =1700
                    Width =2426
                    Height =1235
                    Name ="Box63"
                    LayoutCachedLeft =170
                    LayoutCachedTop =1700
                    LayoutCachedWidth =2596
                    LayoutCachedHeight =2935
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =2655
                    Top =1695
                    Width =2711
                    Height =1235
                    Name ="Box64"
                    LayoutCachedLeft =2655
                    LayoutCachedTop =1695
                    LayoutCachedWidth =5366
                    LayoutCachedHeight =2930
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =225
                    Top =1755
                    Width =2265
                    Height =450
                    Name ="Label65"
                    Caption ="Top Competitors and \015\012Event Entry (T&&F Only)"
                    LayoutCachedLeft =225
                    LayoutCachedTop =1755
                    LayoutCachedWidth =2490
                    LayoutCachedHeight =2205
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =2761
                    Top =1750
                    Width =2445
                    Height =225
                    Name ="Label66"
                    Caption ="Only Top Competitors:"
                    LayoutCachedLeft =2761
                    LayoutCachedTop =1750
                    LayoutCachedWidth =5206
                    LayoutCachedHeight =1975
                End
                Begin Label
                    OverlapFlags =215
                    Left =226
                    Top =1417
                    Width =1987
                    Height =220
                    FontWeight =700
                    Name ="Label67"
                    Caption ="Export Data"
                    LayoutCachedLeft =226
                    LayoutCachedTop =1417
                    LayoutCachedWidth =2213
                    LayoutCachedHeight =1637
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2746
                    Top =1975
                    Width =1155
                    Height =225
                    Name ="Label68"
                    Caption ="Track && Field"
                    LayoutCachedLeft =2746
                    LayoutCachedTop =1975
                    LayoutCachedWidth =3901
                    LayoutCachedHeight =2200
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4006
                    Top =1975
                    Width =1215
                    Height =225
                    Name ="Label69"
                    Caption ="Swimming"
                    LayoutCachedLeft =4006
                    LayoutCachedTop =1975
                    LayoutCachedWidth =5221
                    LayoutCachedHeight =2200
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =226
                    Top =2267
                    Width =1029
                    Height =570
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    Name ="MMmapping"
                    Caption ="Configure\015\012Age Mapping"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =226
                    LayoutCachedTop =2267
                    LayoutCachedWidth =1255
                    LayoutCachedHeight =2837
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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

Private Sub Help_Click()
    ' Need to move focus to entry and then press F1
    Mteam.SetFocus
    Call ShowHelp
End Sub

Private Sub MMExportComp_Click()
    On Error GoTo MMExportComp_Click_Err
    
    Dim strFile As String
    Dim strQuery As String
    
    ' Save record to set Export details
    DoCmd.RunCommand acCmdSaveRecord

    strFile = MMSaveFile
    If strFile <> "" Then
        Call ExportMeetManager("MeetManagerAthletes", strFile)
    End If
   
MMExportComp_Click_Exit:
    Exit Sub

MMExportComp_Click_Err:
    MsgBox ("Error in MMExportComp_Click: " & Error$)
    GoTo MMExportComp_Click_Exit
End Sub

Private Sub MMExportEntry_Click()
    On Error GoTo MMExportEntry_Click_Err
    
    Dim strFile As String
    Dim strQuery As String
   
    ' Save record to set Export details
    DoCmd.RunCommand acCmdSaveRecord
    
    'strQuery = "SELECT * FROM MeetManagerEvents;"
    'Call ExportMMQuery(strQuery)
    
    strFile = MMSaveFile
    If strFile <> "" Then
        Call ExportMeetManager("MeetManagerEvents", strFile)
    End If


MMExportEntry_Click_Exit:
    Exit Sub

MMExportEntry_Click_Err:
    MsgBox ("Error in MMExportEntry_Click: " & Error$)
    GoTo MMExportEntry_Click_Exit
End Sub


Private Function MMSaveFile() As String
    On Error GoTo MMSaveFile_Err
    
    Dim strFilter As String
    Dim strFile As String
    Dim strDefaultFile As String
    Dim lngFlags As Long
    
    Dim strTitle As String
    Dim varFileName As Variant
    
    strTitle = "Meet Manager Export"
    strDefaultFile = "MeetManager-" & Format(Date, "yyyy-mm-dd") & ".txt"
    
    strFilter = ahtAddFilterItem(strFilter, "Text Files (*.txt)", "*.txt")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    lngFlags = ahtOFN_OVERWRITEPROMPT Or ahtOFN_HIDEREADONLY
    
    varFileName = ahtCommonFileOpenSave( _
               OpenFile:=False, _
               Filter:=strFilter, _
               Flags:=lngFlags, _
               FileName:=strDefaultFile, _
               DialogTitle:=strTitle)
               
    If varFileName <> "" Then
        strFile = Trim(varFileName)
    End If
    
MMSaveFile_Exit:
    MMSaveFile = strFile
    Exit Function

MMSaveFile_Err:
    MsgBox ("Error in MMSaveFile: " & Error$)
    GoTo MMSaveFile_Exit
End Function

 Public Sub ExportMMQuery(exportSQL As String)
    Dim db As DAO.Database, qd As DAO.QueryDef
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)

    Set db = CurrentDb

    'Check to see if querydef exists
    For i = 0 To (db.QueryDefs.Count - 1)
        If db.QueryDefs(i).Name = "tmpExport" Then
            db.QueryDefs.Delete ("tmpExport")
            Exit For
    End If
    Next i

    Set qd = db.CreateQueryDef("tmpExport", exportSQL)

    'Set intial filename
    fd.InitialFileName = "MeetManager-" & Format(Date, "yyyy-mm-dd") & ".txt"

    If fd.Show = True Then
        If Format(fd.SelectedItems(1)) <> vbNullString Then
            DoCmd.TransferText acExportDelim, "MMEvents", "tmpExport", fd.SelectedItems(1), False
        End If
    End If
    
    ' NoTextQual  MMEvents

    'Cleanup
    db.QueryDefs.Delete "tmpExport"
    db.Close
    Set db = Nothing
    Set qd = Nothing
    Set fd = Nothing

    End Sub
    
    
Public Function ExportMeetManager(sQuery As String, sFilePath As String)
    On Error GoTo ExportMeetManager_Err
    ' Code based from https://access-programmers.co.uk/forums/showthread.php?t=202570
    
    Dim Rs As DAO.Recordset
    Dim ff As Long
    Dim nIndex As Integer
    Dim sStr As String
    Dim nErrors As Integer
    
    nErrors = 0
    
    Set Rs = CurrentDb.OpenRecordset(sQuery)
    
    ff = FreeFile
    
    Open sFilePath For Output As #ff
     
    Do Until Rs.EOF
        ' Check Data is OK
        If IsError(Rs(0)) Then
            ' Need to add some debugging to  let user know about error.
            nErrors = nErrors + 1
            Rs.MoveNext
        End If
        
        'Queries are single column
        sStr = Trim(Rs(0))
        
        'Write record to the file
        Print #ff, sStr
        
        'Reset the sStr to ""
        sStr = ""
        Rs.MoveNext
    Loop
    
    If nErrors > 0 Then
        MsgBox ("While Export " & nErrors & " Errors were found")
    End If
    
ExportMeetManager_Exit:
    Close #ff
    Rs.Close
    Set Rs = Nothing
    Exit Function

ExportMeetManager_Err:
    MsgBox ("Error in ExportMeetManager: " & Error$)
    GoTo ExportMeetManager_Exit
End Function

Private Sub MMExportRE1_Click()
 On Error GoTo MMExportRE1_Click_Err
    
    Dim strFile As String
    Dim strQuery As String
    Dim lngFlags As Long
    
    ' Save record to set Export details
    DoCmd.RunCommand acCmdSaveRecord

    'strFile = MMSaveFile
    
    Dim strFilter As String
    Dim strDefaultFile As String
    
    Dim strTitle As String
    Dim varFileName As Variant
    
    strTitle = "Meet Manager Registration Export"
    strDefaultFile = "MeetManager-" & Format(Date, "yyyy-mm-dd") & ".re1"
    
    lngFlags = ahtOFN_OVERWRITEPROMPT Or ahtOFN_HIDEREADONLY
    
    strFilter = ahtAddFilterItem(strFilter, "Registration File (*.re1)", "*.re1")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    varFileName = ahtCommonFileOpenSave( _
               OpenFile:=False, _
               Filter:=strFilter, _
               Flags:=lngFlags, _
               FileName:=strDefaultFile, _
               DialogTitle:=strTitle)
               
    If varFileName <> "" Then
        strFile = Trim(varFileName)
    End If
    

    If strFile <> "" Then
        Call ExportToRE1("MeetManagerRE1", strFile)
    End If
   
MMExportRE1_Click_Exit:
    Exit Sub

MMExportRE1_Click_Err:
    MsgBox ("Error in MMExportRE1_Click: " & Error$)
    GoTo MMExportRE1_Click_Exit
End Sub

Public Function ExportToRE1(sQuery As String, sFilePath As String)
    ' Code based from https://access-programmers.co.uk/forums/showthread.php?t=202570
    On Error GoTo ExportToRE1_Err
    
    Dim Rs As DAO.Recordset
    Dim ff As Long
    
    Dim sStr As String
    Dim sTeam As String
    Dim sTAbbr As String
    Dim sID As String
    
    Set Rs = CurrentDb.OpenRecordset(sQuery)
    
    ff = FreeFile
    
    sTeam = Clean(DLookup("[Mteam]", "Miscellaneous"))
    sTAbbr = DLookup("[Mcode]", "Miscellaneous")
    
    Open sFilePath For Output As #ff
    ' Header Line
    ' e.g. "USA Swimming Registration List";"5/28/2009";"LSCDB for Windows";"1.08";"SE"
    sStr = Replace(Clean(DLookup("[CarnivalTitle]", "Miscellaneous")), """", "")
    sStr = """" & sStr & """;""" & Format(Now(), "mm/dd/yyyy") & _
    """;""Sports Administrator"";""" & VersionNumber & """"
    
    '     sStr = """" & DLookup("[CarnivalTitle]", "Miscellaneous") & """;""" & Format(Now(), "mm/dd/yyyy") & _
    """;""Sports Administrator"";""" & VersionNumber & """;""SE"""
    
    Print #ff, sStr
    
    sStr = ""
     
    Do Until Rs.EOF
       '/Loop though all the fields for each record and concat them together with a comma delimiter only.
       '/Note:The Trim(Rs(nIndex) & "") syntax contends with Null or ZLS fields
    
       '"030391JOHFDOEL";"Doel";"John";"J";"F";"03/03/1991";"HURR";"Hurricane Swim Club";"Johnny";"N"
       
       ' ID is PIN or if not set a mix of DoB and name
       
        If Rs("ID") <> "" Then
            sID = Rs("ID")
        Else
            sID = Format(Rs("DOB"), "yymmdd") & Left(Rs("Given"), 3) & Rs("Sex") & Left(Rs("Surname"), 3)
        End If
       
       sStr = """" & sID & """;""" & Rs("Surname") & """;""" & Rs("Given") & """;;""" & Rs("Sex") & """;""" _
       & Format(Rs("DOB"), "mm/dd/yyyy") & """;""" & sTAbbr & """;""" & sTeam & """;""" & Rs("Given") & """;""N"""
    
       '/Write record to the csv file
    
       Print #ff, sStr
    
       '/Reset the sStr to ""
    
       sStr = ""
       Rs.MoveNext
    Loop
    Close #ff
    
ExportToRE1_Exit:
    Rs.Close
    Set Rs = Nothing
    Exit Function

ExportToRE1_Err:
    MsgBox ("Error in ExportToRE1: " & Error$)
    GoTo ExportToRE1_Exit
End Function

Private Sub MMmapping_Click()
    ' Open Form
        Dim DocName As String
    Dim LinkCriteria As String

    DocName = "MeetManagerDivisions"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acWindowNormal

    
End Sub
