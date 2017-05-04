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
    Width =5499
    ItemSuffix =62
    Left =-19995
    Top =2565
    Right =-12735
    Bottom =8430
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
            Height =2948
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =113
                    Top =163
                    Width =4541
                    Height =1984
                    Name ="Box50"
                    LayoutCachedLeft =113
                    LayoutCachedTop =163
                    LayoutCachedWidth =4654
                    LayoutCachedHeight =2147
                End
                Begin Label
                    OverlapFlags =215
                    Left =226
                    Top =226
                    Width =1987
                    Height =220
                    FontWeight =700
                    Name ="Label56"
                    Caption ="Meet Manager Export"
                    LayoutCachedLeft =226
                    LayoutCachedTop =226
                    LayoutCachedWidth =2213
                    LayoutCachedHeight =446
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1302
                    Top =566
                    Width =3021
                    Name ="Mteam"
                    ControlSource ="Mteam"

                    LayoutCachedLeft =1302
                    LayoutCachedTop =566
                    LayoutCachedWidth =4323
                    LayoutCachedHeight =806
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =340
                            Top =570
                            Width =975
                            Height =240
                            Name ="Label57"
                            Caption ="Team Name:"
                            LayoutCachedLeft =340
                            LayoutCachedTop =570
                            LayoutCachedWidth =1315
                            LayoutCachedHeight =810
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1302
                    Top =963
                    Width =606
                    TabIndex =1
                    Name ="Mcode"
                    ControlSource ="Mcode"

                    LayoutCachedLeft =1302
                    LayoutCachedTop =963
                    LayoutCachedWidth =1908
                    LayoutCachedHeight =1203
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =340
                            Top =960
                            Width =945
                            Height =240
                            Name ="Label58"
                            Caption ="Team Code:"
                            LayoutCachedLeft =340
                            LayoutCachedTop =960
                            LayoutCachedWidth =1285
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =3741
                    Top =963
                    Width =516
                    TabIndex =2
                    Name ="Mtop"
                    ControlSource ="Mtop"

                    LayoutCachedLeft =3741
                    LayoutCachedTop =963
                    LayoutCachedWidth =4257
                    LayoutCachedHeight =1203
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2154
                            Top =963
                            Width =1545
                            Height =240
                            Name ="Label59"
                            Caption ="Number Competiors:"
                            LayoutCachedLeft =2154
                            LayoutCachedTop =963
                            LayoutCachedWidth =3699
                            LayoutCachedHeight =1203
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =566
                    Top =1417
                    Width =1584
                    Height =600
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    Name ="MMExportEntry"
                    Caption ="Export Competitors and Entries"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =566
                    LayoutCachedTop =1417
                    LayoutCachedWidth =2150
                    LayoutCachedHeight =2017
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =2607
                    Top =1417
                    Width =1584
                    Height =600
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    Name ="MMExportComp"
                    Caption ="Export Only Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    LayoutCachedLeft =2607
                    LayoutCachedTop =1417
                    LayoutCachedWidth =4191
                    LayoutCachedHeight =2017
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
                    Left =113
                    Top =2324
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

                    LayoutCachedLeft =113
                    LayoutCachedTop =2324
                    LayoutCachedWidth =1247
                    LayoutCachedHeight =2834
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
    
    Dim strTitle As String
    Dim varFileName As Variant
    
    strTitle = "Meet Manager Export"
    strDefaultFile = "MeetManager-" & Format(Date, "yyyy-mm-dd") & ".txt"
    
    strFilter = ahtAddFilterItem(strFilter, "Text Files (*.txt)", "*.txt")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    varFileName = ahtCommonFileOpenSave( _
               OpenFile:=False, _
               Filter:=strFilter, _
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
    
    Set Rs = CurrentDb.OpenRecordset(sQuery)
    
    ff = FreeFile
    
    Open sFilePath For Output As #ff
     
    Do Until Rs.EOF
        'Queries are single column
        sStr = Trim(Rs(0))
        
        'Write record to the file
        Print #ff, sStr
        
        'Reset the sStr to ""
        sStr = ""
        Rs.MoveNext
    Loop
    Close #ff
    
ExportMeetManager_Exit:
    Rs.Close
    Set Rs = Nothing
    Exit Function

ExportMeetManager_Err:
    MsgBox ("Error in ExportMeetManager: " & Error$)
    GoTo ExportMeetManager_Exit
End Function
