Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =9949
    ItemSuffix =78
    Left =-18870
    Top =2760
    Right =-7125
    Bottom =11490
    HelpContextId =80
    RecSrcDt = Begin
        0x98553b042dc7e140
    End
    RecordSource ="MiscellaneousLocal"
    Caption ="Enter Competitors in Events"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
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
        Begin CheckBox
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
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =420
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =90
                    Top =60
                    Width =4695
                    Height =360
                    FontSize =12
                    FontWeight =600
                    Name ="Label77"
                    Caption ="Enter Competitors in Events"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6177
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    ColumnCount =7
                    Left =2168
                    Top =114
                    Width =5875
                    Height =5930
                    TabIndex =13
                    BorderColor =12632256
                    Name ="Summary"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Heats.HE_Code, Heats.E_Number AS [#], EventType.ET_Des AS Event, Events.S"
                        "ex, Events.Age, Trim(Str([F_Lev])) & \" / \" & Trim(Str([Heat])) AS [Final/Heat]"
                        ", IIf([Completed],\"Yes\",\"No\") AS Comp, Heats.Status, Heats.F_Lev, Events.Inc"
                        "lude, EventType.Include FROM EventType INNER JOIN (Events INNER JOIN Heats ON Ev"
                        "ents.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.ET_Code WHERE (((Heats"
                        ".E_Number) Like [forms]![CompEventsSummary]![Event#] Or (Heats.E_Number) Is Null"
                        ") AND ((EventType.ET_Des) Like [forms]![CompEventsSummary]![Event]) AND ((Events"
                        ".Sex)=[forms]![CompEventsSummary]![Male] Or (Events.Sex)=[forms]![CompEventsSumm"
                        "ary]![Female] Or (Events.Sex)=[forms]![CompEventsSummary]![Mixed]) AND ((Events."
                        "Age) Like [forms]![CompEventsSummary]![Age]) AND ((Heats.Status)=[forms]![CompEv"
                        "entsSummary]![Future] Or (Heats.Status)=[forms]![CompEventsSummary]![Active] Or "
                        "(Heats.Status)=[forms]![CompEventsSummary]![Completed] Or (Heats.Status)=[forms]"
                        "![CompEventsSummary]![Promoted]) AND ((Heats.F_Lev) Like [forms]![CompEventsSumm"
                        "ary]![Level]) AND ((Events.Include)=True) AND ((EventType.Include)=True) AND ((H"
                        "eats.Completed) Like [forms]![CompEventsSummary]![HeatComplete])) ORDER BY IIf(I"
                        "sNull([E_Number]),9999999,[E_Number]), EventType.ET_Des, Events.Sex, Events.Age,"
                        " Heats.F_Lev, Val([Heat]);"
                    ColumnWidths ="0;510;2439;626;631;864;567"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnKeyDown ="[Event Procedure]"
                    ControlTipText ="Double-Click an event to manage it (ie. add / remove competitors and enter resul"
                        "ts)."
                    HorizontalAnchor =2
                    VerticalAnchor =2

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8220
                    Top =5645
                    Width =1588
                    Height =397
                    FontSize =8
                    TabIndex =12
                    Name ="CloseBut"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =340
                    Top =356
                    Name ="FutureCB"
                    ControlSource ="CESfuture"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    ControlTipText ="Show finals that will happen in the future."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =600
                            Top =356
                            Width =585
                            Height =240
                            Name ="Text19"
                            Caption ="Future"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =340
                    Top =633
                    TabIndex =1
                    Name ="ActiveCB"
                    ControlSource ="CESactive"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=True"
                    ControlTipText ="Show finals that are currently being contested."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =600
                            Top =633
                            Width =585
                            Height =240
                            Name ="Text22"
                            Caption ="Active"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =340
                    Top =918
                    TabIndex =2
                    Name ="CompletedCB"
                    ControlSource ="CEScompleted"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    ControlTipText ="Show finals that have been completed.  (Only these can be promoted)."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =597
                            Top =916
                            Width =825
                            Height =240
                            Name ="Text24"
                            Caption ="Completed"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =340
                    Top =1201
                    TabIndex =3
                    Name ="PromotedCB"
                    ControlSource ="CESpromoted"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"
                    ControlTipText ="Show finals that have been promoted."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =597
                            Top =1199
                            Width =750
                            Height =240
                            Name ="Text26"
                            Caption ="Promoted"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Top =356
                    Width =231
                    TabIndex =14
                    Name ="Future"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Top =633
                    Width =231
                    TabIndex =15
                    Name ="Active"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Top =916
                    Width =231
                    TabIndex =16
                    Name ="Completed"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Top =1200
                    Width =231
                    TabIndex =17
                    Name ="Promoted"
                    FontName ="Tahoma"

                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =116
                    Top =188
                    Width =1793
                    Height =1866
                    Name ="Box32"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =175
                    Top =72
                    Width =1020
                    Height =210
                    BackColor =-2147483633
                    Name ="Text33"
                    Caption ="Show Finals"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8220
                    Top =1305
                    Width =1588
                    Height =567
                    FontSize =8
                    FontWeight =400
                    TabIndex =10
                    ForeColor =8404992
                    Name ="PromoteBut"
                    Caption ="Promote ALL"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Promote competitors into the next final level for all completed final levels."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =340
                    Top =2444
                    TabIndex =5
                    Name ="MaleCB"
                    ControlSource ="CESmale"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="True"
                    ControlTipText ="Show male events."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =600
                            Top =2444
                            Width =585
                            Height =240
                            Name ="Text36"
                            Caption ="Male"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =340
                    Top =2721
                    TabIndex =6
                    Name ="FemaleCB"
                    ControlSource ="CESfemale"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=True"
                    ControlTipText ="Show female events."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =600
                            Top =2721
                            Width =585
                            Height =240
                            Name ="Text38"
                            Caption ="Female"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Top =2444
                    Width =231
                    TabIndex =18
                    Name ="Male"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Top =2721
                    Width =231
                    TabIndex =19
                    Name ="Female"
                    FontName ="Tahoma"

                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =116
                    Top =2276
                    Width =1793
                    Height =2871
                    Name ="Box47"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =175
                    Top =2160
                    Width =570
                    Height =210
                    BackColor =-2147483633
                    Name ="Text48"
                    Caption ="Show:"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    OverlapFlags =247
                    ListWidth =2268
                    Left =286
                    Top =3980
                    Width =1420
                    Height =223
                    TabIndex =8
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"10\";\"60\""
                    Name ="Event"
                    ControlSource ="CESevent"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT EventType.ET_Des, EventType.Include FROM EventType WHERE ((EventType.Incl"
                        "ude=True)) ORDER BY EventType.ET_Des;"
                    ColumnWidths ="1111"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    FontName ="Tahoma"
                    ControlTipText ="Show the selected event."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =288
                            Top =3760
                            Width =1200
                            Height =240
                            Name ="Text50"
                            Caption ="Event (* for all):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    Left =286
                    Top =3527
                    Width =1420
                    Height =223
                    TabIndex =7
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"10\";\"20\""
                    Name ="Age"
                    ControlSource ="CESage"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Events.Age FROM Events ORDER BY Events.Age;"
                    ColumnWidths ="1111"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    FontName ="Tahoma"
                    ControlTipText ="Show events for the selected age."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =288
                            Top =3307
                            Width =1200
                            Height =240
                            Name ="Text52"
                            Caption ="Age (* for all):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8220
                    Top =1926
                    Width =1588
                    Height =567
                    FontSize =8
                    FontWeight =400
                    TabIndex =11
                    ForeColor =8404992
                    Name ="PromoteSelectedBut"
                    Caption ="Promote SELECTED"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Promote competitors into the next final level for the selected event only."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =247
                    ListWidth =1360
                    Left =284
                    Top =1703
                    Width =1420
                    Height =223
                    TabIndex =4
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"2\";\"1\""
                    Name ="Level"
                    ControlSource ="CESflevel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Heats.F_Lev FROM Heats ORDER BY Heats.F_Lev;"
                    ColumnWidths ="1111"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    FontName ="Tahoma"
                    ControlTipText ="Show a particular final level (* for all)."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =286
                            Top =1483
                            Width =1200
                            Height =240
                            Name ="Text57"
                            Caption ="Level (* for all):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =342
                    Top =4781
                    TabIndex =9
                    Name ="HeatCompleteCB"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="No"
                    ControlTipText ="Show completed heats."

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =606
                            Top =4785
                            Width =1200
                            Height =240
                            Name ="Text61"
                            Caption ="Heat Completed"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Top =4428
                    Width =231
                    TabIndex =20
                    Name ="HeatComplete"
                    DefaultValue ="=\"*\""
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    OverlapFlags =247
                    ListWidth =1360
                    Left =285
                    Top =4434
                    Width =1420
                    Height =223
                    TabIndex =21
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"4\";\"4\""
                    Name ="Event#"
                    ControlSource ="CESevent#"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Heats.E_Number FROM EventType INNER JOIN (Events INNER JOIN Heats ON Even"
                        "ts.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.ET_Code WHERE (((Events."
                        "Include)=Yes) AND ((EventType.Include)=Yes)) ORDER BY Heats.E_Number;"
                    ColumnWidths ="1111"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    FontName ="Tahoma"
                    EventProcPrefix ="Event_"
                    ControlTipText ="Show the event with the selected number."

                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =294
                            Top =4211
                            Width =1335
                            Height =240
                            Name ="Text64"
                            Caption ="Event # (* for all):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8220
                    Top =4965
                    Width =1588
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =22
                    HelpContextId =80
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8220
                    Top =120
                    Width =1588
                    Height =567
                    FontSize =8
                    FontWeight =400
                    TabIndex =23
                    ForeColor =8404992
                    Name ="Edit"
                    Caption ="Manage the Selected Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Manage the competitors in the event selected on the left."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =1
                    OverlapFlags =93
                    Left =2438
                    Top =1588
                    Width =3340
                    Height =1077
                    BackColor =12632256
                    Name ="PWbox"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =2
                    Left =2603
                    Top =1938
                    Width =3017
                    Height =557
                    FontWeight =700
                    BackColor =12632256
                    Name ="PWtext"
                    Caption ="Please wait ..."
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8220
                    Top =3045
                    Width =1588
                    Height =567
                    FontSize =8
                    FontWeight =400
                    TabIndex =24
                    ForeColor =8404992
                    Name ="UpdateFinalStat"
                    Caption ="Update Final Status"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Maintenance Option: Update the status of a final level."
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =150
                    Top =5287
                    TabIndex =25
                    Name ="AlertToRecord"
                    ControlSource ="AlertToRecord"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =415
                            Top =5292
                            Width =1590
                            Height =240
                            Name ="Text69"
                            Caption ="Alert to new records"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =150
                    Top =5652
                    TabIndex =26
                    Name ="NoAllocatedLane"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =415
                            Top =5652
                            Width =1650
                            Height =525
                            Name ="Text73"
                            Caption ="Show \"No allocated lane alert\""
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =345
                    Top =3007
                    TabIndex =27
                    Name ="MixedCB"
                    ControlSource ="CESmixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=True"
                    ControlTipText ="Show mixed  events."

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =605
                            Top =3007
                            Width =585
                            Height =240
                            Name ="Label75"
                            Caption ="Mixed"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Top =3037
                    Width =231
                    TabIndex =28
                    Name ="Mixed"
                    FontName ="Tahoma"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database   'Use database order for string comparisons

' Form Dimensions
Dim lMinHeight As Long
Dim lMinWidth As Long

Private Sub ActiveCB_AfterUpdate()

    If [ActiveCB] Then
        [Active] = 1
    Else
        [Active] = 99
    End If

    [Summary].Requery

End Sub

Private Sub AddBut_Click()
On Error GoTo Err_AddBut_Click


    DoCmd.OpenForm "Competitors", , , , , acDialog, "ADD"
    [Summary].Requery


Exit_AddBut_Click:
    Exit Sub

Err_AddBut_Click:
    MsgBox Error$
    Resume Exit_AddBut_Click
    
End Sub

Private Sub Age_AfterUpdate()

    [Summary].Requery

End Sub

Private Sub Age_DblClick(Cancel As Integer)

    Me![Age] = "*"
    Me![Summary].Requery

End Sub

Private Sub AlertToRecord_AfterUpdate()

On Error GoTo Err_AlertToRecord

    DoCmd.RunCommand acCmdSaveRecord

Exit_AlertToRecord:
    Exit Sub

Err_AlertToRecord:
    MsgBox Error$
    Resume Exit_AlertToRecord

End Sub

Private Sub Button70_Click()
On Error GoTo Err_Button70_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Button70_Click:
    Exit Sub

Err_Button70_Click:
    MsgBox Error$
    Resume Exit_Button70_Click
    
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

Private Sub CompletedCB_AfterUpdate()


    If [CompletedCB] Then
        [Completed] = 2
    Else
        [Completed] = 99
    End If

    [Summary].Requery

End Sub

Private Sub CopyBut_Click()
On Error GoTo Err_CopyBut_Click

    If IsNull([Summary]) Then
        MsgBox ("You must select an event to copy.")
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

Private Sub DeleteBut_Click()
On Error GoTo Err_DeleteBut_Click

    ' Generate Warning - # Competitors, Records,
    
    NumCompEvent = DCount("[PIN]", "CompEvents", "[PIN] = " & [Summary])

    WarningMessage = "This competitor is presently in " & NumCompEvent & " event.  If you continue with this delete operation, all this data will be lost.  Do you wish to continue?"

    Response = MsgBox(WarningMessage, vbCritical + vbYesNo)
        
    If Response = vbYes Then 'Yes
        Q = "DELETE DISTINCTROW Competitors.PIN FROM Competitors WHERE Competitors.PIN= " & [Summary]
        DoCmd.RunSQL Q
        [Summary].Requery
    End If

Exit_DeleteBut_Click:
    Exit Sub

Err_DeleteBut_Click:
    MsgBox Error$
    Resume Exit_DeleteBut_Click
    
End Sub

Private Sub Edit_Click()

  If IsNull(Me!Summary) Then
    Response = MsgBox("Select an event from the list you wish to manage then click the 'Manage the Selected Event' button.", vbOKOnly + vbInformation)
  Else
    Summary_DblClick (Cancel)
  End If

End Sub

Private Sub Event__AfterUpdate()

    Me![Summary].Requery

End Sub

Private Sub Event__DblClick(Cancel As Integer)

    Me![Event#] = "*"
    Me![Summary].Requery

End Sub

Private Sub Event_AfterUpdate()

    [Summary].Requery
    
End Sub

Private Sub Event_DblClick(Cancel As Integer)

    Me![Event] = "*"
    Me![Summary].Requery

End Sub

Private Sub FemaleCB_AfterUpdate()

    If [FemaleCB] Then
        [Female] = "F"
    Else
        [Female] = "X"
    End If

    Me.Summary.Requery

End Sub

Private Sub Field63_AfterUpdate()

    [Summary].Requery

End Sub

Private Sub Form_Load()

    FutureCB_AfterUpdate
    ActiveCB_AfterUpdate
    CompletedCB_AfterUpdate
    PromotedCB_AfterUpdate
    MaleCB_AfterUpdate
    FemaleCB_AfterUpdate

End Sub

Private Sub Form_Open(Cancel As Integer)
    lMinHeight = frmHeight(Me)
    lMinWidth = Me.Width
End Sub

Private Sub Form_Resize()
    If Not m_blResize Then Call glrMinWindowSize(Me, lMinHeight, lMinWidth, True)
End Sub

Private Sub FutureCB_AfterUpdate()

    If FutureCB Then
        [Future] = 0
    Else
        [Future] = 99
    End If

    [Summary].Requery

End Sub

Private Sub HeatCompleteCB_AfterUpdate()

'    If [HeatComplete] = "*" Then
'        [HeatComplete] = True
'        [HeatCompleteCB] = True
'
'    ElseIf Val([HeatComplete]) = -1 Then
'        [HeatComplete] = False
'        [HeatCompleteCB] = False
'
'    Else   ' [HeatComplete] = False
'        [HeatComplete] = "*"
'        [HeatCompleteCB] = Null
'
'    End If

    If [HeatComplete] = True Then
        [HeatComplete] = False
        [HeatCompleteCB] = False

    ElseIf [HeatComplete] = False Then   ' [HeatComplete] = False
        [HeatComplete] = "*"
        [HeatCompleteCB] = Null

    Else    ' [HeatComplete] = Null
        [HeatComplete] = True
        [HeatCompleteCB] = True

    End If
    [Summary].Requery


End Sub

Private Sub Level_AfterUpdate()

    [Summary].Requery

End Sub

Private Sub Level_DblClick(Cancel As Integer)

    Me![Level] = "*"
    Me![Summary].Requery

End Sub

Private Sub MaleCB_AfterUpdate()

    If Me.MaleCB Then
        Me.Male = "M"
    Else
        Me.Male = "X"
    End If

    [Summary].Requery

End Sub

Private Sub MixedCB_AfterUpdate()

  If Me.MixedCB Then
    Me.Mixed = "-"
  Else
    Me.Mixed = "X"
  End If
  Me.Summary.Requery
  
End Sub

Private Sub NoAllocatedLane_AfterUpdate()

    Q = "UPDATE [Misc-EnterCompetitorEvents] SET [ShowNoAllocatedLane]=" & Me!NoAllocatedLane
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True

End Sub

Private Sub PromoteBut_Click()

On Error GoTo Err_PromoteBut_Click
    
    DoCmd.SetWarnings True
    GlobalCancel = False

    Dim Db As Database, rs As Recordset, EventsPromoted As Variant
    
    EventsPromoted = False

    Set Db = CurrentDb()
    
    Q = "SELECT DISTINCT Events.E_Code, Heats.Status, Heats.F_Lev "
    Q = Q & "FROM EventType INNER JOIN (Events INNER JOIN Heats ON Events.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.ET_Code "
    Q = Q & "WHERE Heats.Status=2 and Heats.F_Lev <> 0 "

    Set rs = Db.OpenRecordset(Q, dbOpenDynaset)   ' Create Recordset.

    TotalEvents = rs.RecordCount

    If TotalEvents = 0 Then
        MsgBox ("There are no Finals to be promoted.")
    Else
      'Response = MsgBox("Are you sure you want to promote all finals that have been completed?", 20)
      'If Response = vbYes Then
        
        ReturnValue = SysCmd(acSysCmdInitMeter, "Promoting Competitors", TotalEvents)    ' Display message in status bar.
            
        rs.MoveFirst
        X = 1
        
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE DISTINCTROW ShowDialog SET ShowDialog.ShowDialog = Yes"
        DoCmd.SetWarnings True
    

        While Not rs.EOF And Not GlobalCancel
            'Stop
            If rs!F_Lev <> 0 Then
                E_Code = rs!E_Code
                Ev = DLookup("[ET_Des]", "Events in Full", "[E_Code]=" & E_Code)
                Ev = Ev & "  Age: " & DLookup("[Age]", "Events in Full", "[E_Code]=" & E_Code)
                Ev = Ev & "  Sex: " & DLookup("[Sex]", "Events in Full", "[E_Code]=" & E_Code)
                
                'Response = MsgBox(Message, 20)
                
                GlobalNo = False
                
                If DLookup("[ShowDialog]", "ShowDialog") = True Then
                    DoCmd.OpenForm "PromoteEvents", , , , , acDialog, Ev
                End If
                
                Me.Repaint
                'Stop
                
                If Not GlobalNo And Not GlobalCancel Then
                    
                    Result = PromoteEventFinal(rs!E_Code)
                    If Result = True Then
                        EventsPromoted = True
                    End If
                    ReturnValue = SysCmd(acSysCmdUpdateMeter, X)   ' Update meter.
                    X = X + 1
                End If
            End If
            rs.MoveNext
    
        Wend
        
        ReturnValue = SysCmd(acSysCmdRemoveMeter)
        
        [Summary].Requery
        If EventsPromoted Then
            Response = MsgBox("The events you have just promoted are now set as promoted.", vbInformation)
        Else
            Response = MsgBox("No events were promoted.", vbInformation)
        End If

    End If

    rs.Close




Exit_PromoteBut_Click:
    Set Db = Nothing
    Exit Sub

Err_PromoteBut_Click:
    MsgBox Error$
    Resume Exit_PromoteBut_Click
    
End Sub

Private Sub PromotedCB_AfterUpdate()

    If [PromotedCB] Then
        [Promoted] = 3
    Else
        [Promoted] = 99
    End If

    [Summary].Requery

End Sub

Private Sub PromoteSelectedBut_Click()

    Dim Db As Database, rs As Recordset
    Set Db = DBEngine.Workspaces(0).Databases(0)
    
    Set rs = Db.OpenRecordset("Heats", dbOpenDynaset)   ' Create Recordset.

  If IsNull([Summary]) Then
      MsgBox ("You must select an event from the 'Completed' Final list.")
  Else
    Criteria = "[HE_Code]= " & [Summary]
    rs.FindFirst Criteria
    
    If rs.NoMatch Or rs!Status <> 2 Then
        MsgBox ("You must select an event from the 'Completed' Final list.")
    ElseIf rs!F_Lev = 0 Then
        MsgBox ("The event you are trying to promote is at the highest final level.  There is no final for competitors to be promoted into.")
        
    Else
        E_Code = rs!E_Code
        F_Lev = rs!F_Lev
        E_Des = EventDescription(E_Code)
        E_Sex = EventSex(E_Code)
        E_Age = EventAge(E_Code)
        'Message = "This will promote ALL heats in the final level you have selected.  Are you sure you want to promote this final level:  Final Level=" & F_LEv & "  Age=" & E_Age & "  Sex=" & E_Sex & "  Event=" & E_Des & "?"
        'Stop
        DoCmd.SetWarnings False
        DoCmd.RunSQL "UPDATE DISTINCTROW ShowDialog SET ShowDialog.ShowDialog = Yes"
        DoCmd.SetWarnings True
        
        Ev = DLookup("[ET_Des]", "Events in Full", "[E_Code]=" & E_Code)
        Ev = Ev & "  Age: " & DLookup("[Age]", "Events in Full", "[E_Code]=" & E_Code)
        Ev = Ev & "  Sex: " & DLookup("[Sex]", "Events in Full", "[E_Code]=" & E_Code)
        
        'Response = MsgBox(Message, 20)
        
        GlobalCancel = False
        GlobalNo = False
        
        If DLookup("[ShowDialog]", "ShowDialog") = True Then
            DoCmd.OpenForm "PromoteEvents", , , , , acDialog, Ev
        End If
        
        If Not GlobalNo And Not GlobalCancel Then
'            Stop
            success = PromoteEventFinal(E_Code)
            If success Then
                [Summary].Requery
                MsgBox "The event you have just promoted is now flagged as promoted.  To view it select the 'Promoted' check box in the 'Show Finals' section of this form.", vbInformation
            End If
        End If
    End If
  End If

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

Private Sub ShowPleaseWait()

    If Me![PWbox].visible = True Then
        Me![PWbox].visible = False
        Me![PWtext].visible = False
    Else
        Me![PWbox].visible = True
        Me![PWtext].visible = True
    End If
    DoCmd.RepaintObject A_FORM, "EnterCompetitors"

End Sub

Private Sub Summary_DblClick(Cancel As Integer)

On Error GoTo err_sdc


    If Not IsNull([Summary]) Then
        PleaseWaitMsg = "Retrieving event data ..."
        DoCmd.RunMacro "ShowPleaseWait"
        'DoCmd.OpenForm "EnterCompetitors", , , "[HE_Code] = " & [Summary], , acDialog
        DoCmd.OpenForm "EnterCompetitors", , , "[HE_Code] = " & [Summary], , acWindowNormal
        [Summary].Requery
    End If
                     
exit_sdc:
    Exit Sub

err_sdc:
    MsgBox (Error$)
    GoTo exit_sdc

End Sub

Private Sub Summary_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        Summary_DblClick (Cancel)
    End If

End Sub

Private Sub UpdateFinalStat_Click()
On Error GoTo Err_UpdateFinalStat_Click


    If Not IsNull(Me![Summary]) Then
        vE_Code = DLookup("[E_Code]", "Events in Full", "[HE_Code]=" & Me![Summary])
        Call SetCurrentFinal(vE_Code)
    End If

    Me![Summary].Requery

Exit_UpdateFinalStat_Click:
    Exit Sub

Err_UpdateFinalStat_Click:
    MsgBox Error$
    Resume Exit_UpdateFinalStat_Click
    
End Sub
