Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    WhatsThisButton = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =20
    Width =11469
    ItemSuffix =80
    Left =375
    Top =795
    Right =14100
    Bottom =9480
    HelpContextId =390
    Filter ="[ET_Code] = 3"
    RecSrcDt = Begin
        0xee9763512d29e240
    End
    RecordSource ="SELECT DISTINCTROW EventType.ET_Code, EventType.ET_Des, EventType.Units, EventTy"
        "pe.Lane_Cnt, EventType.R_Code, ReportTypes.Desc, EventType.EntrantNum, EventType"
        ".Include, EventType.PlacesAcrossAllHeats FROM ReportTypes INNER JOIN EventType O"
        "N ReportTypes.R_Code = EventType.R_Code;"
    Caption ="Event Details"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0xa2050000a1050000a1050000a105000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
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
            LabelX =-154
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =187
            Height =187
            LabelX =-154
        End
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Height =255
            LabelX =-154
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            LabelX =-154
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Height =255
            LabelX =-154
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
            Height =6944
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =149
                    Top =2192
                    Width =9988
                    Height =4654
                    BackColor =12632256
                    Name ="Box44"
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =3847
                    Top =2810
                    Width =6190
                    Height =3956
                    BackColor =12632256
                    Name ="Box62"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =2160
                    Top =288
                    Width =3165
                    Height =256
                    BorderColor =12632256
                    HelpContextId =400
                    Name ="ET_Des"
                    ControlSource ="ET_Des"
                    StatusBarText ="Event Description - ie. 200m; 100m Hurdles; High Jump- No sex / age specifics"
                    FontName ="Tahoma"
                    ControlTipText ="A description of the event."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =936
                            Top =288
                            Width =1140
                            Height =245
                            FontWeight =400
                            Name ="Text19"
                            Caption ="Description:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =2157
                    Top =1121
                    Width =930
                    Height =256
                    TabIndex =2
                    BorderColor =12632256
                    HelpContextId =420
                    Name ="Lane_Cnt"
                    ControlSource ="Lane_Cnt"
                    StatusBarText ="Lane / Competitor Count"
                    ValidationRule =">=0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="How many competitors can compete in a single heat."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =288
                            Top =1080
                            Width =1785
                            Height =485
                            FontWeight =400
                            Name ="Text23"
                            Caption ="Lane / Competitor Count (0 if unlimited):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    SpecialEffect =3
                    Left =281
                    Top =2525
                    Width =3430
                    Height =4242
                    TabIndex =5
                    Name ="ET_Sub1"
                    SourceObject ="Form.EventTypeSub1"
                    LinkChildFields ="ET_Code"
                    LinkMasterFields ="ET_Code"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =0
                            Left =255
                            Top =2235
                            Width =1230
                            Height =240
                            BackColor =-2147483633
                            Name ="Text34"
                            Caption ="Divisions"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =7388
                    Top =2520
                    Width =1005
                    Height =225
                    FontSize =6
                    TabIndex =10
                    Name ="Button35"
                    Caption ="REORDER"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Reorder the heats below."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =3402
                    Left =2151
                    Top =720
                    Width =3175
                    Height =256
                    TabIndex =1
                    BorderColor =12632256
                    HelpContextId =410
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="R_Code"
                    ControlSource ="R_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportTypes.R_Code, ReportTypes.Desc, ReportTypes.EventReport FROM Report"
                        "Types WHERE ((ReportTypes.EventReport=Yes)) ORDER BY ReportTypes.Desc;"
                    ColumnWidths ="0;3451"
                    FontName ="Tahoma"
                    ControlTipText ="Select the appropriate report style for this event."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =827
                            Top =720
                            Width =1260
                            Height =245
                            FontWeight =400
                            Name ="Text39"
                            Caption ="Report Style:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    ColumnCount =2
                    ListWidth =1210
                    Left =3743
                    Top =1127
                    Width =1600
                    Height =256
                    TabIndex =3
                    BorderColor =12632256
                    HelpContextId =430
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Units"
                    ControlSource ="Units"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW Units.DisplayUnit, Units.Unit FROM Units;"
                    ColumnWidths ="0;902"
                    FontName ="Tahoma"
                    ControlTipText ="What units are used to measure the results of this event."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =3244
                            Top =1127
                            Width =465
                            Height =245
                            FontWeight =400
                            Name ="Text46"
                            Caption ="Units"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =6040
                    Top =288
                    Width =680
                    Height =238
                    TabIndex =11
                    Name ="ET_Code"
                    ControlSource ="ET_Code"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =10275
                    Top =819
                    Width =1134
                    Height =555
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    HelpContextId =530
                    ForeColor =8404992
                    Name ="SetupHeats"
                    Caption ="Quickly Setup Heats"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Quickly setup heats for all divisions."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Subform
                    OverlapFlags =215
                    SpecialEffect =3
                    Left =3840
                    Top =3207
                    Width =6192
                    Height =3552
                    TabIndex =12
                    Name ="ET_Sub2"
                    SourceObject ="Form.EventTypeSub2"
                    LinkChildFields ="E_Code"
                    LinkMasterFields ="vE_Code"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =4621
                    Top =2923
                    Width =822
                    Height =249
                    FontSize =9
                    FontWeight =700
                    TabIndex =13
                    BackColor =-2147483633
                    Name ="vAge"
                    ControlSource ="=[Forms]![EventType]![ET_Sub1].[Form]![Age]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    Left =6111
                    Top =2923
                    Width =237
                    Height =249
                    FontSize =9
                    FontWeight =700
                    TabIndex =14
                    BackColor =-2147483633
                    Name ="vSex"
                    ControlSource ="=[Forms]![EventType]![ET_Sub1].[Form]![Sex]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =5832
                    Top =595
                    Width =702
                    TabIndex =15
                    Name ="vE_Code"
                    ControlSource ="=[Forms]![EventType]![ET_Sub1].[Form]![E_Code]"
                    FontName ="Tahoma"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    Left =3827
                    Top =2917
                    Width =480
                    Height =270
                    FontSize =9
                    FontWeight =400
                    Name ="Text58"
                    Caption ="Age:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    Left =5545
                    Top =2925
                    Width =465
                    Height =255
                    FontSize =9
                    FontWeight =400
                    Name ="Text59"
                    Caption ="Sex:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =3862
                    Top =2520
                    Width =2760
                    Height =225
                    Name ="Text60"
                    Caption ="Heats for selected division:"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =10275
                    Top =6330
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =9
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =144
                    Top =144
                    Width =9994
                    Height =1874
                    Name ="Box63"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =10275
                    Top =195
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    HelpContextId =520
                    ForeColor =8404992
                    Name ="Lane Allocation"
                    Caption ="Lane Promotion"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    EventProcPrefix ="Lane_Allocation"
                    ControlTipText ="Setup what lanes competitors receive when promoted into the next final level."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =4752
                    Top =1584
                    Width =585
                    Height =245
                    TabIndex =4
                    BorderColor =12632256
                    HelpContextId =440
                    Name ="EntrantNum"
                    ControlSource ="EntrantNum"
                    StatusBarText ="Lane / Competitor Count"
                    FontName ="Tahoma"
                    ControlTipText ="This is required only if you are generating carnival disks."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =247
                            Left =717
                            Top =1641
                            Width =3915
                            Height =245
                            FontWeight =400
                            Name ="Text66"
                            Caption ="Number of Entrants from each House/School:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =10275
                    Top =5595
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    HelpContextId =60
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =8855
                    Top =1052
                    TabIndex =16
                    BorderColor =12632256
                    Name ="Include"
                    ControlSource ="Include"
                    DefaultValue ="Yes"
                    ControlTipText ="Tick if you want this event included in the carnival."

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6825
                            Top =1050
                            Width =1935
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text72"
                            Caption ="Include Event in Carnival:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =8742
                    Top =2880
                    TabIndex =17
                    HelpContextId =495
                    BorderColor =12632256
                    Name ="Sync"
                    DefaultValue ="Yes"
                    ControlTipText ="Untick this to set different pointscales within a final-level."

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6792
                            Top =2880
                            Width =1875
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text74"
                            Caption ="Sychronize Point Scales:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =8855
                    Top =1412
                    TabIndex =18
                    BorderColor =12632256
                    Name ="Check78"
                    ControlSource ="PlacesAcrossAllHeats"
                    DefaultValue ="Yes"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6120
                            Top =1410
                            Width =2640
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Label79"
                            Caption ="Places determined across all heats:"
                            FontName ="Tahoma"
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

Private Sub Button29_Click()
On Error GoTo Err_Button29_Click


    DoCmd.GoToRecord , , A_NEWREC

Exit_Button29_Click:
    Exit Sub

Err_Button29_Click:
    MsgBox Error$
    Resume Exit_Button29_Click
    
End Sub

Private Sub Button30_Click()
On Error GoTo Err_Button30_Click


    DoCmd.RunCommand acCmdFind

Exit_Button30_Click:
    Exit Sub

Err_Button30_Click:
    MsgBox Error$
    Resume Exit_Button30_Click
    
End Sub

Private Sub Button31_Click()
On Error GoTo Err_Button31_Click


    DoCmd.GoToRecord , , A_NEXT

Exit_Button31_Click:
    Exit Sub

Err_Button31_Click:
    MsgBox Error$
    Resume Exit_Button31_Click
    
End Sub

Private Sub Button32_Click()
On Error GoTo Err_Button32_Click


    DoCmd.GoToRecord , , A_PREVIOUS

Exit_Button32_Click:
    Exit Sub

Err_Button32_Click:
    MsgBox Error$
    Resume Exit_Button32_Click
    
End Sub

Private Sub Button35_Click()
    
    Forms![EventType]![ET_Sub1].Form.Requery
    Forms![EventType]![ET_Sub2].Form.Requery


End Sub

Private Sub Button36_Click()
On Error GoTo Err_Button36_Click


    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdCopy
    DoCmd.RunCommand acCmdPasteAppend 'Paste Append

Exit_Button36_Click:
    Exit Sub

Err_Button36_Click:
    MsgBox Error$
    Resume Exit_Button36_Click
    
End Sub

Private Sub Button42_Click()
On Error GoTo Err_Button42_Click


    DoCmd.GoToRecord , , A_PREVIOUS

Exit_Button42_Click:
    Exit Sub

Err_Button42_Click:
    MsgBox Error$
    Resume Exit_Button42_Click
    
End Sub

Private Sub Button43_Click()
On Error GoTo Err_Button43_Click


    DoCmd.GoToRecord , , A_NEXT

Exit_Button43_Click:
    Exit Sub

Err_Button43_Click:
    MsgBox Error$
    Resume Exit_Button43_Click
    
End Sub

Private Sub Lane_Cnt_BeforeUpdate(Cancel As Integer)

  If Not IsNumeric(Me!Lane_Cnt) Then
    Response = MsgBox("The Lane Count must be a number.", vbInformation)
    Cancel = True
  End If
    
End Sub

Private Sub SetupHeats_Click()
On Error GoTo Err_SetupHeats_Click
    
    'Dim PrevControl As Control
    'Set PrevControl = Screen.PreviousControl

    'If DCount("[PtScale]", "PointsScale") = 0 Then
    '    DoCmd.GoToControl PrevControl.Name
    '    MsgBox ("You have not entered any points scales. You must have entered at least one point scale before setting up a heat.")
    'Else

        Dim DocName As String
        Dim LinkCriteria As String

        DocName = "Final_lev"
        LinkCriteria = "[ET_Code] = Forms![EventType]![ET_Code]"
        DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog
        Me![ET_Sub1].Requery
        Me![ET_Sub2].Requery

    'End If

Exit_SetupHeats_Click:
    Exit Sub

Err_SetupHeats_Click:
    MsgBox Error$
    Resume Exit_SetupHeats_Click
    
End Sub

Private Sub Button68_Click()

End Sub

Private Sub Button75_Click()
On Error GoTo Err_Button75_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Button75_Click:
    Exit Sub

Err_Button75_Click:
    MsgBox Error$
    Resume Exit_Button75_Click
    
End Sub

Private Sub Cancel_Click()

On Error GoTo Err_Cancel_Click

    'DoCmd.RunCommand acCmdUndo

Exit_Cancel_Click:
    DoCmd.Close
    Exit Sub

Err_Cancel_Click:
'    MsgBox Error$
    Resume Exit_Cancel_Click
    
    
End Sub

Private Sub Close_Click()
On Error GoTo Err_Close_But_Click

    'Stop
    Cl = True

    X = DCount("[ET_Code]", "Lane Promotion Allocation", "[ET_Code]=" & Me![ET_Code])

    If X < Forms![EventType]![Lane_Cnt] And DCount("[ET_Code]", "EventTypeHeats", "[ET_Code]=" & Me![ET_Code] & " AND [F_Lev]>0") Then
        Response = MsgBox("The number of promotion lanes you have set up is less than the Lane / Competitor count.  Do you still wish to continue?", vbYesNo + vbQuestion, "Too few lanes?")
        If Response <> vbYes Then
            Cl = False
        End If
    End If

    
    If IsNull([R_Code]) Then
        [R_Code] = -1
    End If

    If IsNull(DLookup("[R_Code]", "ReportTypes", "[R_Code]=" & Me![R_Code])) And (Cl = True) Then
        MsgBox ("You must choose a Report Style for the event.")
        Cl = False
    End If

    If Me![Lane_Cnt] = 0 Then
        If Not IsNull(DLookup("[F_Lev]", "Events in Full", "[F_Lev]>0 AND [ET_Code]=" & Me![ET_Code])) Then
            For i = 0 To (DMax("[F_Lev]", "Events in Full", "[F_Lev]>0 AND [ET_Code]=" & Me![ET_Code]) - 1)
                ProNum = DLookup("[ProNum]", "Final_Lev", "[ET_Code]=" & Me![ET_Code] & " AND [F_Lev]=" & i)
                Response = vbYes
                If IsNull(ProNum) Or ProNum = 0 Then
                    Response = MsgBox("You have not entered the number of competitors that are to be promoted into final level " & i & ".  This is required to automatically promote the best competitors.  It is set in the Setup Heats form.  Do you wish to continue?", vbYesNo + vbCritical, "Promotion Details Incomplete")
                End If
                If Response = vbNo Then
                    Cl = False
                    GoTo Exit_Close_But_Click
                End If

                
            Next i

        End If
    End If

    If Cl Then
        DoCmd.Close
    End If

Exit_Close_But_Click:
    Exit Sub

Err_Close_But_Click:
    MsgBox Error$
    Resume Exit_Close_But_Click
    


End Sub

Private Sub ET_Sub2_Enter()

    'Stop

    Dim PrevControl As control
    Set PrevControl = Screen.PreviousControl
    
    If IsNull([vE_Code]) Then
        DoCmd.GoToControl PrevControl.Name
        Response = MsgBox("You must select (or create) a 'Division' before modifying the heats for the division.", vbInformation)
    ElseIf DCount("[PtScale]", "PointsScale") = 0 Then
        DoCmd.GoToControl PrevControl.Name
        Response = MsgBox("You have not setup any points-scales. You must have setup at least one point-scale before you can setup up a heat.", vbInformation)

    End If

End Sub

Private Sub ET_Sub2_Exit(Cancel As Integer)

    If Not (IsNull([vE_Code])) Then
        Call SetCurrentFinal([vE_Code])
    End If

    If Not CheckFinalIntegrity(vE_Code, "HEATS") Then
         Response = MsgBox("Finals should be in consecutive order starting at 0 and increasing by one (1) only.  This is not necessary but is recommended.  Do you wish to continue?", vbYesNo + vbCritical, "Final Integrity Warning")
         If Response = vbYes Then
            Cancel = False
         Else
            Cancel = True
         End If
    End If


End Sub

Private Sub Form_Load()

    If Me.[OpenArgs] = "ADD" Then
        Me.[DefaultEditing] = 1
        DoCmd.GoToRecord A_FORM, "EventType", A_NEWREC
    Else
        Me.[DefaultEditing] = 4
    End If

    

End Sub

Private Sub Lane_Allocation_Click()

On Error GoTo Err_LA

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "Lane Promotion Allocation"
    LinkCriteria = "[ET_Code] = Forms![EventType]![ET_Code]"
    DoCmd.OpenForm DocName, , , , , acDialog

Exit_LA:
    Exit Sub

Err_LA:
    MsgBox Error$
    Resume Exit_LA
    


End Sub

Private Sub Lane_Cnt_AfterUpdate()
  
On Error GoTo Lane_Cnt_AfterUpdate_Err
  
  Call UpdateLaneTemplate(Me!ET_Code, Me!Lane_Cnt)
  
Lane_Cnt_AfterUpdate_Exit:
  Exit Sub
  
Lane_Cnt_AfterUpdate_Err:
  MsgBox ("An error has occured in [EventTypeWizard:Lane_Cnt_AfterUpdate]: " & Err.Description)
  GoTo Lane_Cnt_AfterUpdate_Exit
  
End Sub
