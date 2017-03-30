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
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridY =10
    Width =7143
    ItemSuffix =123
    Left =-22905
    Top =5370
    Right =-15795
    Bottom =10305
    RecSrcDt = Begin
        0xb3dbca3c8df5e140
    End
    RecordSource ="Misc-EventLists"
    Caption ="Utilities 2"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
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
        Begin OptionButton
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =187
            Height =187
            LabelX =-236
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =187
            Height =187
            LabelX =-236
        End
        Begin TextBox
            OldBorderStyle =0
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
            BackColor =12632256
        End
        Begin ListBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
            BackColor =12632256
        End
        Begin ComboBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
            BackColor =12632256
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =0
            BackColor =128
            Name ="FormHeader1"
        End
        Begin Section
            CanGrow = NotDefault
            Height =6435
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =170
                    Top =113
                    Width =5508
                    Height =4581
                    BackColor =11589887
                    Name ="Box51"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    ColumnCount =2
                    ListWidth =1375
                    Left =3628
                    Top =680
                    Width =1525
                    Height =227
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="Sex_DD"
                    ControlSource ="Rsex"
                    RowSourceType ="Value List"
                    RowSource ="\"*\";\"Any\";\"M\";\"Male\";\"F\";\"Female\""
                    ColumnWidths ="390;735"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Left =3001
                            Top =684
                            Width =420
                            Height =240
                            FontWeight =400
                            Name ="Sex_DD_Tit"
                            Caption ="Sex"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5904
                    Top =144
                    Width =1134
                    Height =600
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="RemoveEmpty"
                    Caption ="Remove Empty Heats"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =215
                    ListWidth =1510
                    Left =3625
                    Top =1021
                    Width =1540
                    Height =225
                    TabIndex =2
                    BackColor =16777215
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"2\";\"1\""
                    Name ="Flev_DD"
                    ControlSource ="Rfinal"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Heats.F_Lev FROM Heats ORDER BY Heats.F_Lev;"
                    ColumnWidths ="1510"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Left =2774
                            Top =1021
                            Width =630
                            Height =240
                            FontWeight =400
                            Name ="Text71"
                            Caption ="Final"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =223
                    ListWidth =1510
                    Left =3640
                    Top =1588
                    Width =1510
                    Height =225
                    TabIndex =3
                    BackColor =16777215
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="Heat_EB"
                    ControlSource ="Rheat"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Heats.Heat FROM Heats ORDER BY Heats.Heat;"
                    ColumnWidths ="1510"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =223
                            Left =2774
                            Top =1588
                            Width =630
                            Height =240
                            FontWeight =400
                            Name ="Text73"
                            Caption ="Heat"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    ListWidth =1510
                    Left =3628
                    Top =340
                    Width =1525
                    Height =225
                    TabIndex =4
                    BackColor =16777215
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"10\";\"20\""
                    Name ="Age_EB"
                    ControlSource ="Rage"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Events.Age FROM Events;"
                    ColumnWidths ="1510"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Left =2777
                            Top =340
                            Width =630
                            Height =240
                            FontWeight =400
                            Name ="Text75"
                            Caption ="Age"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =336
                    Top =340
                    Width =2138
                    Height =2346
                    TabIndex =5
                    BorderColor =12632256
                    Name ="EventSF"
                    SourceObject ="Form.Report SF2"

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =3660
                    Top =2041
                    TabIndex =6
                    BorderColor =12632256
                    Name ="Future"
                    ControlSource ="Rfuture"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =223
                            Left =2944
                            Top =1984
                            Width =615
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text81"
                            Caption ="Future"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =3660
                    Top =2325
                    TabIndex =7
                    BorderColor =12632256
                    Name ="Active"
                    ControlSource ="Ractive"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =223
                            Left =2944
                            Top =2268
                            Width =615
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text83"
                            Caption ="Active"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =4966
                    Top =2041
                    TabIndex =8
                    BorderColor =12632256
                    Name ="Completed"
                    ControlSource ="Rcompleted"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =223
                            Left =3965
                            Top =1984
                            Width =900
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text85"
                            Caption ="Completed"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =223
                    Left =4966
                    Top =2325
                    TabIndex =9
                    BorderColor =12632256
                    Name ="Promoted"
                    ControlSource ="Rpromoted"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =223
                            Left =3965
                            Top =2268
                            Width =900
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text87"
                            Caption ="Promoted"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    BackStyle =0
                    Left =3850
                    Top =1984
                    Width =216
                    TabIndex =10
                    BackColor =-2147483633
                    Name ="FutureEB"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    BackStyle =0
                    Left =3850
                    Top =2267
                    Width =216
                    TabIndex =11
                    BackColor =-2147483633
                    Name ="ActiveEB"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    BackStyle =0
                    Left =5269
                    Top =1984
                    Width =216
                    TabIndex =12
                    BackColor =-2147483633
                    Name ="CompletedEB"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    BackStyle =0
                    Left =5269
                    Top =2267
                    Width =216
                    TabIndex =13
                    BackColor =-2147483633
                    Name ="PromotedEB"

                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =247
                    Left =2661
                    Top =1417
                    Width =2778
                    Height =1248
                    Name ="Box93"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5865
                    Top =3165
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =14
                    HelpContextId =410
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
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5865
                    Top =4015
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =15
                    Name ="Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    Left =1318
                    Top =5045
                    Width =4131
                    TabIndex =16
                    BackColor =-2147483633
                    Name ="Checking"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =85
                            Left =283
                            Top =5045
                            Width =945
                            Height =240
                            BackColor =-2147483633
                            Name ="Text120"
                            Caption ="Checking:"
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    OldBorderStyle =1
                    OverlapFlags =215
                    BackStyle =0
                    Left =1330
                    Top =2948
                    Width =4110
                    Height =1575
                    TabIndex =17
                    BackColor =-2147483633
                    BorderColor =12632256
                    Name ="Changes"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =215
                            Left =340
                            Top =2948
                            Width =915
                            Height =240
                            BackColor =-2147483633
                            Name ="Text122"
                            Caption ="Changes:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter2"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons

Private Sub Active_AfterUpdate()


    If Me![Active] = True Then
        Me![ActiveEB] = 1
    Else
        Me![ActiveEB] = 9
    End If



End Sub

Private Sub Age_EB_DblClick(Cancel As Integer)

    Me![Age_EB] = "*"

End Sub

Private Sub Button112_Click()
On Error GoTo Err_Button112_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Button112_Click:
    Exit Sub

Err_Button112_Click:
    MsgBox Error$
    Resume Exit_Button112_Click
    
End Sub

Private Sub Button38_Click()

End Sub

Private Sub Button47_Click()

End Sub

Private Sub Button64_Click()
On Error GoTo Err_Button64_Click


    DoCmd.Close

Exit_Button64_Click:
    Exit Sub

Err_Button64_Click:
    MsgBox Error$
    Resume Exit_Button64_Click
    
End Sub

Private Sub Button65_Click()
On Error GoTo Err_Button65_Click
     
    DoCmd.RunCommand acCmdSaveRecord

    Dim Criteria As String, Db As Database, rs As Recordset
    Dim NewTitle As String
    Set Db = CurrentDb()

    Q = "SELECT DISTINCTROW EventType.ET_Code, EventType.Flag, EventType.R_Code "
    Q = Q & "FROM EventType WHERE EventType.Flag = True ORDER BY EventType.R_Code"

    Set rs = Db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.
    
    rs.MoveFirst
    
    Old_R_Code = -1

    Do Until rs.EOF  ' Loop until no matching records.
        
        R_Code = rs!R_Code
        If R_Code <> Old_R_Code Then
            If Me![SummaryReport] Then
                ReportName = DLookup("[SummaryReport]", "ReportTypes", "[R_Code] = " & R_Code)
                If Not IsNull(ReportName) Then
                    DoCmd.OpenReport ReportName, A_PREVIEW, , "[R_Code] = " & R_Code
                End If
            End If
            
            If Me![Detailed] Then
                ReportName = DLookup("[Report]", "ReportTypes", "[R_Code] = " & R_Code)
                If Not IsNull(ReportName) Then
                    DoCmd.OpenReport ReportName, A_PREVIEW, , "[R_Code] = " & R_Code
                End If
            End If
            
            If Me![EntrySheet] Then
                'ReportName = DLookup("[SummaryReport]", "ReportTypes")
                DoCmd.OpenReport "EventResultsEntrySheet", A_PREVIEW
            End If
            
            'If Me![EventResults] Then
            '
            '    DoCmd OpenReport "Results-by Event", A_PREVIEW
            'End If

        End If

        Old_R_Code = R_Code

        rs.MoveNext
    Loop
    
    rs.Close

Exit_Button65_Click:
    Set Db = Nothing
    Exit Sub

Err_Button65_Click:
    MsgBox Error$
    Resume Exit_Button65_Click
    
End Sub

Private Sub Close_Click()

    DoCmd.Close

End Sub

Private Sub Completed_AfterUpdate()

    If Me![Completed] = True Then
        Me![CompletedEB] = 2
    Else
        Me![CompletedEB] = 9
    End If

End Sub

Private Sub Event_DD_DblClick(Cancel As Integer)

    Me![Event_DD] = "*"

End Sub

Private Sub Event_DD_GotFocus()
    
    'If Forms![reports_event]![RSfld] = "Results Entry Sheets" Then
    '    Q = "SELECT EventType.ET_Des, ReportTypes.Desc FROM ReportTypes INNER JOIN EventType ON ReportTypes.R_Code = EventType.R_Code WHERE ReportTypes.Desc like ""*"" ORDER BY EventType.ET_Des;"
    'Else
    '    Q = "SELECT EventType.ET_Des, ReportTypes.Desc FROM ReportTypes INNER JOIN EventType ON ReportTypes.R_Code = EventType.R_Code WHERE ReportTypes.Desc=[forms]![reports_event]![RSfld] ORDER BY EventType.ET_Des;"
    'End If
    '
    'Forms![reports_event]![Event_DD].RowSource = Q
    'Forms![reports_event]![Event_DD].Requery


End Sub

Private Sub Flev_DD_DblClick(Cancel As Integer)

    Me![Flev_DD] = "*"

End Sub

Private Sub Form_Load()

    Future_AfterUpdate
    Active_AfterUpdate
    Completed_AfterUpdate
    Promoted_AfterUpdate

End Sub

Private Sub Form_Open(Cancel As Integer)

    ' Update the Total Points field in Competitor's Table
    '
    'For Each Competitor
    '   Total Points =
    '           Find each event competior is in
    '           Determine point scale and place for each event
    '           Find points allocated to place and add to total points


End Sub

Private Sub Future_AfterUpdate()

    If Me![Future] = True Then
        Me![FutureEB] = 0
    Else
        Me![FutureEB] = 9
    End If


End Sub

Private Sub Heat_EB_DblClick(Cancel As Integer)

    Me![Heat_EB] = "*"

End Sub

Private Sub Promoted_AfterUpdate()

    If Me![Promoted] = True Then
        Me![PromotedEB] = 3
    Else
        Me![PromotedEB] = 9
    End If


End Sub

Private Sub RemoveEmpty_Click()

'Stop
On Error GoTo Err_RemoveEmpty_Click
         
    Response = MsgBox("This action will remove all heats that satisfy your selection and that have NO competitors in.  Care must be taken that future events that have not yet had competitors promoted into them are not accidentally removed.  Do you wish to continue?", 20, "Remove Empty Heats?")
    If Response = 6 Then
        DoCmd.RunCommand acCmdSaveRecord
    
        Dim Criteria As String, Db As Database, rs As Recordset
        Dim NewTitle As String, F_LevCriteria  As String
        
        Set Db = CurrentDb()
    
        Q = "SELECT DISTINCTROW EventType.ET_Code, EventType.Flag, EventType.R_Code "
        Q = Q & "FROM EventType WHERE EventType.Flag = True ORDER BY EventType.R_Code"
    
        Set rs = Db.OpenRecordset("Events in Full", dbOpenDynaset)   ' Create dynaset.
            
        If Me![Flev_DD] = "*" Then
            F_LevCriteria = "like ""*"""
        Else
            F_LevCriteria = "= " & Me![Flev_DD]
        End If


        Criteria = "[Age] like """ & Me![Age_EB] & """ and [Sex] like """ & Me![Sex_DD] & """ and [F_Lev] " & F_LevCriteria & " and [Heat] like """ & Me![Heat_EB] & """ and [Flag]=TRUE"
        Criteria = Criteria & " AND ([Status]=" & Me![FutureEB] & " OR [Status]=" & Me![ActiveEB] & " OR [Status]=" & Me![CompletedEB] & " OR [Status]=" & Me![PromotedEB] & ")"
        
        TotalRecs = DCount("[E_Code]", "Events in Full", Criteria)
        ReturnValue = SysCmd(acSysCmdInitMeter, "Removing empty heats ... ", TotalRecs)    ' Display message in status bar.
        X = 0

        rs.FindFirst Criteria
        
        Old_R_Code = -1
        'Stop
        Debug.Print "+==============================+"
        Do Until rs.EOF Or rs.NoMatch  ' Loop until no matching records.
            ReturnValue = SysCmd(acSysCmdUpdateMeter, X)   ' Update meter.
            X = X + 1
            He = rs!HE_Code
            Crit2 = "[E_Code]=" & rs!E_Code & " and [Heat]=" & rs!Heat & " and [F_Lev]=" & rs!F_Lev
            If IsNull(DLookup("[E_Code]", "CompEvents", Crit2)) Then
                'Stop
                'Debug.Print He, rs.ET_Des, rs.Sex, rs.Age, rs.F_Lev, rs.Heat
                Q = "DELETE DISTINCTROW Heats!E_Code, Heats!Heat, Heats!F_Lev FROM Heats "
                Me![Changes] = Me![Changes] & "Event " & rs!ET_Des & ", " & rs!Sex & ", " & rs!Age & ", " & rs!F_Lev & ", " & rs!Heat & Chr$(13)
                Q = Q & "WHERE " & Crit2
                DoCmd.SetWarnings False
                DoCmd.RunSQL Q
                DoCmd.SetWarnings True
    
            End If
            rs.FindNext Criteria
        Loop
        Debug.Print "+==============================+"
        rs.Close
        ReturnValue = SysCmd(acSysCmdRemoveMeter)
    End If

Exit_RemoveEmpty_Click:
    Set Db = Nothing
    Exit Sub

Err_RemoveEmpty_Click:
    MsgBox Error$
    Resume Exit_RemoveEmpty_Click
    

End Sub

Private Sub Retrieve_Click()

End Sub

Private Sub RSfld_AfterUpdate()
    

    Forms![reports_event]![Event_DD] = "*"


End Sub

Private Sub Sel_Event_Change()

    Y = 1
    z = 2
    K = 3

    X = [Forms]![EnterCompetitors]![Sel_Event]

End Sub

Private Sub Sel_Event_DblClick(Cancel As Integer)


    Y = 1
    z = 2
    K = 3

    X = [Forms]![EnterCompetitors]![Sel_Event]
    Y = X + 1


End Sub

Private Sub Sex_DD_DblClick(Cancel As Integer)

    Me![Sex_DD] = "*"

End Sub
