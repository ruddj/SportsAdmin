Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    GridX =20
    GridY =20
    Width =6632
    ItemSuffix =43
    Left =1935
    Top =585
    Right =11520
    Bottom =7590
    HelpContextId =50
    PaintPalette = Begin
        0x000301000000000000000000
    End
    RecSrcDt = Begin
        0xa89ea4290fcde140
    End
    Caption ="PointsScale"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
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
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
        End
        Begin ListBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
        End
        Begin ComboBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Tab
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin Page
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
            Height =5520
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =5385
                    Top =4904
                    Width =1134
                    Height =510
                    FontSize =8
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5393
                    Top =225
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    ForeColor =32768
                    Name ="Add"
                    Caption ="Add Pointscale"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5393
                    Top =855
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    ForeColor =128
                    Name ="Delete"
                    Caption ="Delete Pointscale"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5393
                    Top =1485
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =8404992
                    Name ="Rename"
                    Caption ="Rename Pointscale"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5393
                    Top =2115
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    ForeColor =8404992
                    Name ="Update"
                    Caption ="Update all Points"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Update all competitor results to reflect the current pointscales."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =5392
                    Top =3223
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    HelpContextId =50
                    Name ="HelpBut"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =90
                    Top =90
                    Width =5175
                    Height =5325
                    TabIndex =6
                    Name ="TabCtl39"
                    FontName ="Tahoma"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =225
                            Top =495
                            Width =4905
                            Height =4785
                            Name ="Page40"
                            Caption ="Pointscales"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    Left =2398
                                    Top =1099
                                    Width =2571
                                    Height =4086
                                    Name ="PointScaleSubform"
                                    SourceObject ="Form.PointScaleSubform"
                                    LinkChildFields ="PtScale"
                                    LinkMasterFields ="PtScale"

                                End
                                Begin ListBox
                                    SpecialEffect =3
                                    OverlapFlags =247
                                    Left =300
                                    Top =797
                                    Width =1690
                                    Height =2270
                                    TabIndex =1
                                    Name ="PtScale"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCT PointsScale.PtScale FROM PointsScale;"
                                    ColumnWidths ="1441"
                                    AfterUpdate ="[Event Procedure]"
                                    OnDblClick ="[Event Procedure]"
                                    FontName ="Tahoma"
                                    OnKeyDown ="[Event Procedure]"

                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =409
                                    Top =571
                                    Width =1530
                                    Height =225
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Text21"
                                    Caption ="Point Scale"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =2285
                                    Top =570
                                    Width =1530
                                    Height =225
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Text22"
                                    Caption ="Allocated Points"
                                    FontName ="Tahoma"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =223
                                    Left =2236
                                    Top =793
                                    Width =2868
                                    Height =4475
                                    Name ="Box24"
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =2806
                                    Top =853
                                    Width =795
                                    Height =225
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Text25"
                                    Caption ="Place"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =1
                                    Left =3589
                                    Top =853
                                    Width =795
                                    Height =225
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Text26"
                                    Caption ="Points"
                                    FontName ="Tahoma"
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    Left =414
                                    Top =3689
                                    Width =801
                                    TabIndex =2
                                    BorderColor =12632256
                                    Name ="NumPlaces"
                                    DefaultValue ="50"
                                    FontName ="Tahoma"
                                    ControlTipText ="Number of places to create quickly."

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =425
                                            Top =3461
                                            Width =1395
                                            Height =225
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text30"
                                            Caption ="Number of Places"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =223
                                    Left =415
                                    Top =4213
                                    Width =801
                                    TabIndex =3
                                    BorderColor =12632256
                                    Name ="NumPoints"
                                    DefaultValue ="1"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =426
                                            Top =3985
                                            Width =1395
                                            Height =225
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text32"
                                            Caption ="Number of Points"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =255
                                    Left =300
                                    Top =3351
                                    Width =1698
                                    Height =1910
                                    Name ="Box33"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =345
                                    Top =3180
                                    Width =1605
                                    Height =225
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Text34"
                                    Caption ="Create Points Quickly"
                                    FontName ="Tahoma"
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontFamily =34
                                    Left =408
                                    Top =4607
                                    Width =1509
                                    Height =510
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =4
                                    Name ="AllocateDefPoints"
                                    Caption ="Create Points"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =225
                            Top =495
                            Width =4905
                            Height =4785
                            Name ="Page41"
                            Caption ="Information"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =340
                                    Top =623
                                    Width =4252
                                    Height =4422
                                    FontWeight =400
                                    Name ="Label42"
                                    Caption ="Each event must have a pointscale allocated to it in order for the Sports Admini"
                                        "strator to determine how many points each place should get.  \015\012\015\012Set"
                                        "up as many pointscales as you require.  Normally a separate pointscale is setup "
                                        "for finals, heats and relays.\015\012\015\012It is important to ensure that poin"
                                        "ts are allocated to every possible place.  For example, if there are 35 competit"
                                        "ors in a single race then the pointscale for that event must have at least 35 pl"
                                        "aces setup.\015\012\015\012This can take a little while to do manually so use th"
                                        "e \"Create Points Quickly\" box to automatically create the specified number of "
                                        "places and assign them the specified number of points.  Note: No existing places"
                                        " and points will be over-written."
                                    FontName ="Tahoma"
                                End
                            End
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

Private Sub Add_Click()
On Error GoTo Err_Add_Click

    DoCmd.OpenForm "PointScale Add", , , , , acDialog
    Me![PtScale].Requery
    Me![PtScale] = DLookup("[PtScale]", "PointsScale")
    Call ShowSubForm

Exit_Add_Click:
    Exit Sub

Err_Add_Click:
    MsgBox Error$
    Resume Exit_Add_Click
    
End Sub

Private Sub AllocateDefPoints_Click()

On Error GoTo AllocateDefPoints_Click_Err
    

    If IsNull(Me![NumPlaces]) Then
        MsgBox ("Please specify how many places you wish to allocate default point to.")
    ElseIf Me![NumPlaces] = 0 Then
        MsgBox ("The number of places must be greater than 0")
    ElseIf IsNull(Me![NumPoints]) Then
        MsgBox ("Please specify how many points you wish to allocate to each place.")
    Else
        'Stop
        DoCmd.Hourglass True
        Dim Criteria As String, db As Database, Rs As Recordset
        
        Set db = DBEngine.Workspaces(0).Databases(0)
        Set Rs = db.OpenRecordset("PointsScale", dbOpenDynaset)   ' Create dynaset.
        
        msg = "Adding default points ..."
        CountRecs = Me![NumPlaces]
        ReturnValue = SysCmd(acSysCmdInitMeter, msg, Me![NumPlaces].Value)    ' Display message in status bar.

        For i = 1 To Me![NumPlaces]
            ReturnValue = SysCmd(acSysCmdUpdateMeter, i)   ' Update meter.
            Criteria = "[PtScale] = """ & Me![PtScale] & """ AND [Place]= " & i
            Rs.FindNext Criteria    ' Find first occurrence.
            If Rs.NoMatch Then
                Rs.AddNew          ' Enable editing.
                Rs!PtScale = Me![PtScale]
                Rs!Place = i
                Rs!Points = Me![NumPoints]
                Rs.Update
            End If

        Next i
        Rs.Close
        Me.Refresh
        ReturnValue = SysCmd(acSysCmdRemoveMeter)   ' Update meter.
        DoCmd.Hourglass False
    End If
    
    
AllocateDefPoints_Click_Exit:
    Exit Sub

AllocateDefPoints_Click_Err:
    MsgBox ("Error in AllocateDefPoints_Click: " & Error$)
    DoCmd.Hourglass False
    ReturnValue = SysCmd(acSysCmdRemoveMeter)   ' Update meter.
    GoTo AllocateDefPoints_Click_Err
    
End Sub

Private Sub Button15_Click()
On Error GoTo Err_Button15_Click


    DoCmd.Close

Exit_Button15_Click:
    Exit Sub

Err_Button15_Click:
    MsgBox Error$
    Resume Exit_Button15_Click
    
End Sub

Private Sub Button19_Click()

End Sub

Private Sub Close_Click()

    DoCmd.Close

End Sub

Private Sub Delete_Click()

    
    p = Me![PtScale]

    If Not IsNull(p) Then
        If IsNull(DLookup("[HE_Code]", "Heats", "[PtScale]=""" & p & """")) Then
            Response = MsgBox("Are you sure you want to delete this pointscale?", vbYesNo + vbCritical, "Delete Pointscale")
            If Response = vbYes Then
                Q = "DELETE DISTINCTROW PointsScale.PtScale FROM PointsScale "
                Q = Q & "WHERE PointsScale.PtScale=""" & p & """"
            
                DoCmd.SetWarnings False
                DoCmd.RunSQL Q
                DoCmd.SetWarnings True
    
                Me![PtScale].Requery
                Me![PtScale] = DLookup("[PtScale]", "PointsScale")
                ShowSubForm
            End If
    
        Else
            MsgBox ("The PointScale that you are trying to delete is currently used by an event.  First change the Event's pointscale and then delete the unused pointscale.")
        End If
    Else
        MsgBox ("You must select a pointscale before attempting to delete it.")
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)

    'Me![PtScale] = DLookup("[PtScale]", "PointsScale")
    ShowSubForm

End Sub

Private Sub PtScale_AfterUpdate()
    
    ShowSubForm

End Sub

Private Sub PtScale_DblClick(Cancel As Integer)

    If IsNull(Me![PtScale]) Then
        MsgBox ("You must select a Pointscale before renaming it.")
    Else
        DoCmd.OpenForm "PointScale Add", , , , , acDialog, "RENAME"
        Me![PtScale].Requery
    End If
        
End Sub

Private Sub PtScale_KeyDown(KeyCode As Integer, Shift As Integer)

    'Stop
    If KeyCode = vbKeyDelete Then
        Delete_Click
    End If

End Sub

Private Sub Rename_Click()
    
    PtScale_DblClick (Cancel)
    ShowSubForm

End Sub

Private Sub ShowSubForm()

    If IsNull(Me![PtScale]) Then
        Me![PointScaleSubform].visible = False
        Me![AllocateDefPoints].visible = False
    Else
        Me![PointScaleSubform].visible = True
        Me![AllocateDefPoints].visible = True
    End If

End Sub

Private Sub Update_Click()

    Response = MsgBox("This will update ALL points allocated to competitors to the current PointScale values.  This action should only be necessary if the points allocated to a place has been changed AND if events have already been completed.  Do you wish to continue?", vbYesNo + vbCritical, "Update Points")
    If Response = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.OpenQuery "Update Competitor Points"
        DoCmd.SetWarnings True
    End If

End Sub
