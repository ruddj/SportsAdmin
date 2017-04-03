Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =7310
    ItemSuffix =42
    Left =5220
    Top =6300
    Right =11520
    Bottom =7590
    RecSrcDt = Begin
        0x64e50da4704ae240
    End
    RecordSource ="SELECT DISTINCTROW Heats.E_Code, Heats.F_Lev AS FinalLevel, Heats.PtScale AS Poi"
        "ntScale, Heats.Heat, Heats.Pro_Type AS PromotionType, Heats.UseTimes, Heats.Comp"
        "leted, Heats.HE_Code, Heats.EffectsRecords FROM Heats ORDER BY Heats.F_Lev, Heat"
        "s.Heat;"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =240
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =750
                    Width =333
                    Height =227
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text10"
                    Caption ="Heat"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1253
                    Width =1234
                    Height =227
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text11"
                    Caption ="Pt. Scale"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2388
                    Width =1365
                    Height =225
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text28"
                    Caption ="Prom. Ty."
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3810
                    Width =559
                    Height =227
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text29"
                    Caption ="Times?"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =408
                    Height =220
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text30"
                    Caption ="Final"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =4365
                    Width =675
                    Height =225
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Label40"
                    Caption ="Records?"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            Height =290
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =720
                    Top =15
                    Width =453
                    Height =245
                    ColumnWidth =825
                    ColumnOrder =2
                    TabIndex =1
                    Name ="Heat"
                    ControlSource ="Heat"
                    StatusBarText ="How many heats are in this final level."
                    ValidationRule =">0 And <1000"
                    ValidationText ="A heat must be between 0 and 1000."
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Each heat in a particular final-level needs a number (1,2,3 etc)"

                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    ListWidth =1690
                    Left =1215
                    Top =15
                    Width =1270
                    Height =245
                    ColumnWidth =1230
                    ColumnOrder =3
                    TabIndex =2
                    ColumnInfo ="\"\";\">\";\"10\";\"20\""
                    Name ="Point Scale"
                    ControlSource ="PointScale"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PointsScale.PtScale FROM PointsScale GROUP BY PointsScale.PtScale;"
                    ColumnWidths ="1440"
                    StatusBarText ="What pointscale is used to allocate points."
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    EventProcPrefix ="Point_Scale"
                    ControlTipText ="What pointscale is used to allocate points."

                End
                Begin CheckBox
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =3990
                    Top =45
                    Width =197
                    Height =245
                    TabIndex =4
                    Name ="UseTimes"
                    ControlSource ="UseTimes"
                    DefaultValue ="Yes"
                    ControlTipText ="Tick if you want to use the competitors results (rather than place) to promote b"
                        "y."

                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =2755
                    Left =2538
                    Top =15
                    Width =1375
                    Height =245
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="Pro_Type"
                    ControlSource ="PromotionType"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Promotion.ProType, Promotion.Desc FROM Promotion;"
                    ColumnWidths ="885;1620"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="How are competitors promoted into the next final level"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =5772
                    Width =741
                    Height =245
                    TabIndex =5
                    Name ="E_Code"
                    ControlSource ="E_Code"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =6569
                    Width =741
                    Height =245
                    TabIndex =6
                    Name ="HE_Code"
                    ControlSource ="HE_Code"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4755
                    Width =441
                    Height =290
                    TabIndex =7
                    Name ="LookAtEvent"
                    Caption ="Command37"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad7007adadadd70000000000007aa07700777700770dd07707777770770a ,
                        0xa07707877770770dd07707e87770770aa0ff00777700ff0dd0fff000000fff0a ,
                        0xa00000000000000dda00d70ff07a00daadadad7007adadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View competitors in the heat."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ColumnCount =2
                    ListWidth =2325
                    Left =45
                    Top =15
                    Width =635
                    Height =246
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"2\";\"1\""
                    Name ="F_Lev"
                    ControlSource ="FinalLevel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Final Level Sub].F_Lev, [Final Level Sub].F_Lev_Sub FROM [Final Level Su"
                        "b] ORDER BY [Final Level Sub].F_Lev;"
                    ColumnWidths ="567;1701"
                    StatusBarText ="Final Levels start at 0 (0: Grand Final, 1:Semi-Final, 2: Quarter-Final etc)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Final Levels start at 0 (0: Grand Final, 1:Semi-Final, 2: Quarter-Final etc)"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5247
                    Width =336
                    Height =290
                    TabIndex =8
                    Name ="DeleteBut"
                    Caption ="Command39"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadadadadadaadad00adad00adaddadad00ad00adada ,
                        0xadadad0000adadaddadadad00adadadaadadad0000adadaddadad00ad00adada ,
                        0xadad00adad00adaddadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Heat."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =4455
                    Top =45
                    Width =197
                    Height =245
                    TabIndex =9
                    Name ="EffectsRecords"
                    ControlSource ="EffectsRecords"
                    DefaultValue ="Yes"
                    ControlTipText ="Tick if you want this heat to effect event records."

                End
            End
        End
        Begin FormFooter
            Height =283
            BackColor =-2147483633
            Name ="FormFooter2"
            Begin
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =85
                    Left =90
                    Top =30
                    Width =4305
                    Height =210
                    FontWeight =700
                    ForeColor =255
                    Name ="NoRecords"
                    Caption ="ALERT: No heats are setup for this division."
                    FontName ="Tahoma"
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

Private Sub Button16_Click()
On Error GoTo Err_Button16_Click


    DoCmd.GoToRecord , , A_NEXT

Exit_Button16_Click:
    Exit Sub

Err_Button16_Click:
    MsgBox Error$
    Resume Exit_Button16_Click
    
End Sub

Private Sub Button17_Click()
On Error GoTo Err_Button17_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Button17_Click:
    Exit Sub

Err_Button17_Click:
    MsgBox Error$
    Resume Exit_Button17_Click
    
End Sub

Private Sub F_Lev_AfterUpdate()

On Error Resume Next

  If Me!F_Lev = 0 Then
    Me!Pro_Type = "NONE"
  Else
    Me!Pro_Type = "Staggered"
  End If
  
End Sub

Private Sub Form_AfterUpdate()

    
    If Forms![EventType]![Sync] = True Then
        Q = "UPDATE DISTINCTROW Heats SET Heats.PtScale = """ & Me![Point Scale] & """"
        Q = Q & " WHERE Heats.E_Code= " & Me![E_Code] & " AND Heats.F_Lev= " & Me![F_Lev]

        q1 = "UPDATE DISTINCTROW Heats SET Heats.Pro_Type = """ & Me![Pro_Type] & """"
        q1 = q1 & " WHERE Heats.E_Code= " & Me![E_Code] & " AND Heats.F_Lev= " & Me![F_Lev]

        q2 = "UPDATE DISTINCTROW Heats SET Heats.UseTimes =" & Me![UseTimes]
        q2 = q2 & " WHERE Heats.E_Code= " & Me![E_Code] & " AND Heats.F_Lev= " & Me![F_Lev]

        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.RunSQL q1
        DoCmd.RunSQL q2
        DoCmd.SetWarnings True
    End If

End Sub

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)

    Dim MyDb As Database

    Response = MsgBox("Deleting this heat will remove any competitors that may be enrolled in the heat.  Do you want to continue?", 17)
    
    If Response = 2 Then
        Cancel = 1
    Else
    
        Set MyDb = DBEngine.Workspaces(0).Databases(0)
      
        Q = "DELETE DISTINCTROW Heats.E_Code, Heats.Heat "
        Q = Q & "FROM Heats "
        Q = Q & "WHERE ((Heats.E_Code=" & Forms![EventType]![ET_Sub].Form![E_Code]
        Q = Q & ") AND (Heats.Heat= " & Forms![EventType]![ET_Sub].Form![Heat]
        Q = Q & "))"

        MyDb.Execute (Q)
        
        
    End If
    
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

    
  If IsNull(Me![Pro_Type]) Then
      Response = MsgBox("You must enter a valid promotion type (or push Esc to cancel).", vbInformation)
      Cancel = True
  ElseIf IsNull(Me![Point Scale]) Then
      Response = MsgBox("You must enter a valid Point Scale (or push Esc to cancel).", vbInformation)
      Cancel = True
  ElseIf IsNull(Me![Heat]) Then
      Response = MsgBox("You must enter the heat number (or push Esc to cancel).", vbInformation)
      Cancel = True
'  ElseIf DCount("[He_Code]", "Heats", "[E_Code]=" & Me![E_Code] & " AND [F_Lev]=" & Me![F_Lev] & " AND [Heat]=" & Me![Heat]) > 0 Then
'      Response = MsgBox("The Final-Level / Heat pair must be unique.  That is, you can't have two heats in the same final-level with the same number.  You need to change the heat number (or push Esc to cancel).", vbInformation)
'      Cancel = True
  End If

End Sub

Private Sub Form_Current()

  If NoFormRecords(Me.RecordsetClone) Then
    Me!NoRecords.visible = True
  Else
    Me!NoRecords.visible = False
  End If
  
End Sub

Private Sub Form_Delete(Cancel As Integer)

    CompInEvent = CompetitorsInEvent(Me![E_Code], Me![F_Lev], Me![Heat])

    If CompInEvent > 0 Then
        Response = MsgBox("There are " & Str(CompInEvent) & " competitor(s) in this heat.  Are you sure you want to delete it?", 20, "Delete Heat")
        If Response = 7 Then Cancel = True
    End If

End Sub

Private Sub Heat_AfterUpdate()

On Error Resume Next
  
  If IsNull(Me!Pro_Type) Then
    If Me!F_Lev = 0 Then
      Me!Pro_Type = "NONE"
    Else
      Me!Pro_Type = "Staggered"
    End If
  End If

End Sub

Private Sub Heat_BeforeUpdate(Cancel As Integer)

  If IsNull(Me![Heat]) Then
    Response = MsgBox("You must enter a Heat (a whole number between 1 and 1000) or push Esc to cancel.", vbInformation)
    Cancel = True
  End If

End Sub

Private Sub LookAtEvent_Click()
On Error GoTo Err_LookAtEvent_Click
      
  DoCmd.RunCommand acCmdSaveRecord
  
  If Not IsNull(Me!HE_Code) Then
      PleaseWaitMsg = "Retrieving event data ..."
      DoCmd.RunMacro "ShowPleaseWait"
      DoCmd.OpenForm "EnterCompetitors", , , "[HE_Code] = " & Me!HE_Code, , acDialog
  End If
                   

Exit_LookAtEvent_Click:
    Exit Sub

Err_LookAtEvent_Click:
  If Err.Number <> 2501 Then
    MsgBox Err.Description
    Resume Exit_LookAtEvent_Click
  End If
  
End Sub
Private Sub DeleteBut_Click()
On Error GoTo Err_DeleteBut_Click

  Response = MsgBox("Are you sure you want to delete this heat?", vbYesNo + vbInformation, "Delete Confirmation")
  If Response = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
  End If

Exit_DeleteBut_Click:
    Exit Sub

Err_DeleteBut_Click:
    MsgBox Err.Description
    Resume Exit_DeleteBut_Click
    
End Sub

Private Sub Point_Scale_BeforeUpdate(Cancel As Integer)

  If IsNull(Me![Point Scale]) Then
    Response = MsgBox("You must select a Point-Scale from the list (or push Esc to cancel).", vbInformation)
    Cancel = True
  End If
  
End Sub

Private Sub Pro_Type_BeforeUpdate(Cancel As Integer)

  If IsNull(Me![Pro_Type]) Then
    Response = MsgBox("You must select a Promotion Type from the list (or push Esc to cancel).", vbInformation)
    Cancel = True
  End If

End Sub
