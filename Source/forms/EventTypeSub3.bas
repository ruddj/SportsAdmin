Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridY =10
    Width =5782
    ItemSuffix =36
    Left =1920
    Right =11160
    Bottom =5415
    RecSrcDt = Begin
        0x6de96f5db6f2e140
    End
    RecordSource ="SELECT DISTINCT Heats.E_Code, Heats.F_Lev AS FinalLevel, Heats.PtScale AS PointS"
        "cale, Heats.Pro_Type AS PromotionType, Heats.UseTimes, Heats.Completed FROM Heat"
        "s ORDER BY Heats.F_Lev;"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
                    OverlapFlags =93
                    TextAlign =2
                    Left =375
                    Width =1279
                    Height =227
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text11"
                    Caption ="Pt. Scale"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1780
                    Width =1410
                    Height =225
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text28"
                    Caption ="Prom. Ty."
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3184
                    Width =514
                    Height =227
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text29"
                    Caption ="Times?"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Width =408
                    Height =220
                    FontSize =7
                    BackColor =-2147483633
                    Name ="Text30"
                    Caption ="Final"
                    FontName ="Arial"
                End
            End
        End
        Begin Section
            Height =238
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2
                    Width =350
                    Height =227
                    Name ="F_Lev"
                    ControlSource ="FinalLevel"
                    FontName ="Arial"

                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    ListWidth =1690
                    Left =375
                    Width =1405
                    Height =227
                    ColumnWidth =1230
                    ColumnOrder =3
                    TabIndex =1
                    ColumnInfo ="\"\";\">\";\"10\";\"20\""
                    Name ="Point Scale"
                    ControlSource ="PointScale"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT PointsScale.PtScale FROM PointsScale GROUP BY PointsScale.PtScale;"
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Point_Scale"

                End
                Begin CheckBox
                    OverlapFlags =85
                    OldBorderStyle =0
                    Left =3420
                    Top =45
                    Width =152
                    Height =151
                    TabIndex =3
                    Name ="UseTimes"
                    ControlSource ="UseTimes"
                    DefaultValue ="Yes"

                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    ColumnCount =2
                    ListWidth =2755
                    Left =1780
                    Width =1540
                    Height =226
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="Pro_Type"
                    ControlSource ="PromotionType"
                    RowSourceType ="Table/Query"
                    RowSource ="Select [ProType],[Desc] From [Promotion];"
                    ColumnWidths ="885;1620"

                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =3719
                    Width =741
                    Height =225
                    TabIndex =4
                    Name ="E_Code"
                    ControlSource ="E_Code"

                End
            End
        End
        Begin FormFooter
            Height =0
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

Private Sub Form_Delete(Cancel As Integer)

    CompInEvent = CompetitorsInEvent(Me![E_Code], Me![F_Lev], Me![Heat])

    If CompInEvent > 0 Then
        Response = MsgBox("There are " & Str(CompInEvent) & " competitor(s) in this heat.  Are you sure you want to delete it?", 20, "Delete Heat")
        If Response = 7 Then Cancel = True
    End If

End Sub
