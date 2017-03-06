Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    ScrollBars =0
    BorderStyle =3
    GridX =20
    GridY =20
    Width =9648
    ItemSuffix =19
    Left =1380
    Top =1545
    Right =11025
    Bottom =8310
    HelpContextId =90
    RecSrcDt = Begin
        0xd614db87edc6e140
    End
    Caption ="Competitor Summary"
    OnClose ="[Event Procedure]"
    HelpFile ="sports.hlp"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
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
        Begin CheckBox
            BorderLineStyle =0
        End
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin Section
            Height =6774
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin ListBox
                    OverlapFlags =85
                    ColumnCount =5
                    Left =216
                    Top =358
                    Width =7465
                    Height =6305
                    BorderColor =12632256
                    Name ="Summary"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Competitors.PIN, UCase([Surname]) & \", \" & [Gname] AS Name, Competitors"
                        ".Age, Competitors.Sex, House.H_NAme AS Team FROM House INNER JOIN Competitors ON"
                        " House.H_Code = Competitors.H_Code WHERE (((House.Include)=Yes)) ORDER BY UCase("
                        "[Surname]) & \", \" & [Gname], Competitors.Age, Competitors.Sex, House.H_NAme;"
                    ColumnWidths ="0;3402;567;567;1701"
                    OnDblClick ="[Event Procedure]"
                    OnKeyDown ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8070
                    Top =6138
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =1
                    Name ="CloseBut"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =448
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    ForeColor =128
                    Name ="DeleteBut"
                    Caption ="Delete Competitor"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =1071
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =32768
                    Name ="AddBut"
                    Caption ="Add Competitor"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =216
                    Top =75
                    Width =1020
                    Height =210
                    FontWeight =700
                    Name ="Text7"
                    Caption ="Name"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =3616
                    Top =72
                    Width =465
                    Height =225
                    FontWeight =700
                    Name ="Text8"
                    Caption ="Age"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =4801
                    Top =72
                    Width =1290
                    Height =225
                    FontWeight =700
                    Name ="Text9"
                    Caption ="Team"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =4186
                    Top =72
                    Width =450
                    Height =225
                    FontWeight =700
                    Name ="Text10"
                    Caption ="Sex"
                    FontName ="Arial"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7993
                    Top =3259
                    TabIndex =4
                    BorderColor =12632256
                    Name ="Show"
                    DefaultValue ="0"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =85
                            Left =8253
                            Top =3260
                            Width =1050
                            Height =525
                            Name ="Text13"
                            Caption ="Show All Competitors"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =5514
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    HelpContextId =70
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =1995
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    ForeColor =8404992
                    Name ="Roll Over"
                    Caption ="Competitor Roll Over"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Roll_Over"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =2625
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    ForeColor =8404992
                    Name ="Roll Back"
                    Caption ="Competitor Roll Back"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Roll_Back"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =3825
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    ForeColor =8404992
                    Name ="Bulk"
                    Caption ="Bulk Maintenance"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8070
                    Top =4455
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =9
                    ForeColor =32768
                    Name ="CreateTeamNames"
                    Caption ="Create Team Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

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

Dim UpdateCompetitorsOrdered

Private Sub AddBut_Click()
On Error GoTo Err_AddBut_Click

    Call MaintainCompetitor("ADD", 0)
    UpdateCompetitorsOrdered = True
    Me!Summary.Requery

Exit_AddBut_Click:
    Exit Sub

Err_AddBut_Click:
    MsgBox Error$
    Resume Exit_AddBut_Click
    
End Sub

Private Sub Bulk_Click()

On Error GoTo Err_Bulk_Click

    DoCmd.OpenForm "Competitors Bulk Maintain", , , , , acDialog, "ADD"
    Me![Summary].Requery
    UpdateCompetitorsOrdered = True

Exit_Bulk_Click:
    Exit Sub

Err_Bulk_Click:
    MsgBox Error$
    Resume Exit_Bulk_Click
    

End Sub

Private Sub Button16_Click()
On Error GoTo Err_Button16_Click

    Dim DocName As String

    DocName = "Roll Competitors Back"
    DoCmd.OpenQuery DocName, A_NORMAL, A_EDIT

Exit_Button16_Click:
    Exit Sub

Err_Button16_Click:
    MsgBox Error$
    Resume Exit_Button16_Click
    
End Sub

Private Sub CloseBut_Click()
On Error GoTo Err_CloseBut_Click
    
    
    DoCmd.Close

Exit_CloseBut_Click:
    Exit Sub

Err_CloseBut_Click:
    MsgBox Error$
    Resume Exit_CloseBut_Click
    
End Sub

Private Sub CopyBut_Click()
On Error GoTo Err_CopyBut_Click

    If IsNull([Summary]) Then
        Response = MsgBox("You must select an event to copy.", vbInformation)
    Else
        'DoCmd OpenForm "EventTypeCopy", , , , , acDialog
        [Summary].Requery

    End If

Exit_CopyBut_Click:
    Exit Sub

Err_CopyBut_Click:
    MsgBox Error$
    Resume Exit_CopyBut_Click
    
End Sub

Private Sub CreateTeamNames_Click()
'On Error GoTo CreateTeamNames_Click_Err

    PleaseWaitMsg = "Adding Team Competitor Names ..."
    DoCmd.RunMacro "ShowPleaseWait"

    Dim Criteria As String, Db As Database, Hrs As Recordset, Ars As Recordset
    Dim NewTitle As String

    Call UpdateEventCompetitorAge
    
    'for each house do
    '   for each age division do
    '       for each sex
    '           if competitor does not exist
    '             Add competitor

    Set Db = DBEngine.Workspaces(0).Databases(0)
    
    Q = " SELECT DISTINCTROW House.H_Code, House.CompPool, House.Include, House.H_ID FROM House "
    Q = Q & "GROUP BY House.H_Code, House.CompPool, House.Include, House.H_ID "
    Q = Q & "HAVING House.Include=Yes"

    Set Hrs = Db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.

    
    Q = "SELECT DISTINCT CompetitorEventAge.Eage, Str(Val([Eage])) AS Age "
    Q = Q & "FROM CompetitorEventAge "
    Q = Q & "GROUP BY CompetitorEventAge.Eage, Str(Val([Eage]))"

    Set Ars = Db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.
    
    If Hrs.BOF Then
      ' No Teams so do nothing
      Response = MsgBox("No teams have been set up so no team competitors can be created.", vbInformation)
    ElseIf Ars.BOF Then
      Response = MsgBox("No event ages have been defined yet.  This is probably due to no competitors existing in the database.", vbInformation)
    Else
      Hrs.MoveFirst
    
      'stop
  
      Do Until Hrs.EOF  ' Loop until no matching records.
          Ars.MoveFirst
          Do Until Ars.EOF
              
              Criteria = "[Include]=Yes and [Gname] = ""Team"" and [Surname]=""" & Hrs!H_Code & """ and [Sex]=""M"" and [Age]=" & Ars!Age
  
              If IsNull(DLookup("[PIN]", "Competitors", Criteria)) Then
                  UpdateCompetitorsOrdered = True
                  DOB = DetermineDOB(Ars!Age)
                  Q = "INSERT INTO Competitors ([Include], [Gname], [Surname], [Sex], [H_Code], [H_ID], [DOB], [Age]) "
                  Q = Q & "VALUES (Yes, ""Team"", """ & Hrs!H_Code & """, ""M"", """ & Hrs!H_Code & """, " & Hrs!H_ID & ", #" & DOB & "#," & Ars!Age & ")"
                  DoCmd.SetWarnings False
                  DoCmd.RunSQL Q
                  DoCmd.SetWarnings True
  
              End If
  
              Criteria = "[Include]=Yes and [Gname] = ""Team"" and [Surname]=""" & Hrs!H_Code & """ and [Sex]=""F"" and [Age]=" & Ars!Age
  
              If IsNull(DLookup("[PIN]", "Competitors", Criteria)) Then
                  UpdateCompetitorsOrdered = True
                  DOB = DetermineDOB(Ars!Age)
                  Q = "INSERT INTO Competitors ([Include], [Gname], [Surname], [Sex], [H_Code], [H_ID], [DOB], [Age]) "
                  Q = Q & "VALUES (Yes, ""Team"", """ & Hrs!H_Code & """, ""F"", """ & Hrs!H_Code & """, " & Hrs!H_ID & ", #" & DOB & "#," & Ars!Age & ")"
                  DoCmd.SetWarnings False
                  DoCmd.RunSQL Q
                  DoCmd.SetWarnings True
  
              End If
              
              Ars.MoveNext
          Loop
          Hrs.MoveNext
          
      Loop
    
    End If
    Ars.Close
    Hrs.Close

    Me![Summary].Requery


CreateTeamNames_Click_Exit:
    DoCmd.SetWarnings True
    DoCmd.RunMacro "ClosePleaseWait"
    Exit Sub

CreateTeamNames_Click_Err:
    MsgBox (Error$)
    GoTo CreateTeamNames_Click_Exit

End Sub

Private Sub DeleteBut_Click()
On Error GoTo Err_DeleteBut_Click

    ' Generate Warning - # Competitors, Records,
    
    NumCompEvent = DCount("[PIN]", "CompEvents", "[PIN] = " & [Summary])

    WarningMessage = "This competitor is presently enrolled in " & NumCompEvent & " event(s).  Do you wish to continue?"

    Response = MsgBox(WarningMessage, vbYesNo + vbCritical + vbDefaultButton2)
        
    If Response = vbYes Then 'Yes
        DoCmd.SetWarnings False
        Q = "DELETE DISTINCTROW Competitors.PIN FROM Competitors WHERE Competitors.PIN= " & [Summary]
        DoCmd.RunSQL Q
'        Q = "DELETE DISTINCTROW CompetitorsOrdered.PIN FROM CompetitorsOrdered WHERE CompetitorsOrdered.PIN= " & [Summary]
'        DoCmd.RunSQL Q
        Call TransferToCompetitorOrdered
        DoCmd.SetWarnings True
        [Summary].Requery
    End If

Exit_DeleteBut_Click:
    Exit Sub

Err_DeleteBut_Click:
    MsgBox Error$
    Resume Exit_DeleteBut_Click
    
End Sub

Private Sub Form_Close()

  Call UpdateEventCompetitorAge
  
    'If UpdateCompetitorsOrdered Then
    '   Call TransferToCompetitorOrdered
    'End If

End Sub

Private Sub Form_Load()

    UpdateCompetitorsOrdered = False

End Sub


Private Sub Refresh_Click()
On Error GoTo Err_Refresh_Click


    DoCmd.RunCommand acCmdRefresh

Exit_Refresh_Click:
    Exit Sub

Err_Refresh_Click:
    MsgBox Error$
    Resume Exit_Refresh_Click
    
End Sub

Private Sub Roll_Back_Click()

On Error GoTo Err_Roll_Back_Click

  Response = MsgBox("Are you sure you want to roll competitors back?", vbYesNo + vbInformation + vbDefaultButton2, "Roll Back Confirmation.")
  If Response = vbYes Then
    Dim DocName As String

    DoCmd.SetWarnings False

    DocName = "Roll Competitors Back"
    DoCmd.OpenQuery DocName, A_NORMAL, A_EDIT
    
    DoCmd.SetWarnings True
    
    Me![Summary].Requery

    UpdateCompetitorsOrdered = True
  End If
  
Exit_Roll_Back_Click:
    Exit Sub

Err_Roll_Back_Click:
    MsgBox Error$
    Resume Exit_Roll_Back_Click
    

End Sub

Private Sub Roll_Over_Click()

On Error GoTo Err_Roll_Over_Click

  Response = MsgBox("Are you sure you want to roll competitors over?", vbYesNo + vbInformation + vbDefaultButton2, "Roll Over Confirmation.")
  If Response = vbYes Then
    Dim DocName As String
    
    DoCmd.SetWarnings False
    
    DocName = "Roll Competitors Over"
    DoCmd.OpenQuery DocName, A_NORMAL, A_EDIT
    
    DoCmd.SetWarnings True

    Me![Summary].Requery

    UpdateCompetitorsOrdered = True
  End If

Exit_Roll_Over_Click:
    Exit Sub

Err_Roll_Over_Click:
    MsgBox Error$
    Resume Exit_Roll_Over_Click
    
End Sub

Private Sub Summary_DblClick(Cancel As Integer)
    
    On Error GoTo Summary_DblClick_Err
    
    If IsNull(Me!Summary) Then
        MsgBox ("You must select a competitor first.")
    Else
        Call MaintainCompetitor("EDIT", Me!Summary)
        Me!Summary.Requery
    End If
Summary_DblClick_Exit:
    Exit Sub
    
Summary_DblClick_Err:
    MsgBox (Error$)
    GoTo Summary_DblClick_Exit
    
End Sub

Private Sub Summary_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 46 Then
        DeleteBut_Click
    ElseIf KeyCode = 13 Then
        Summary_DblClick (Cancel)
    End If

End Sub
