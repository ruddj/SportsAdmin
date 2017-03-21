Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    AllowUpdating =1
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =20
    Width =10080
    ItemSuffix =73
    Left =120
    Top =450
    Right =12045
    Bottom =8760
    HelpContextId =70
    RecSrcDt = Begin
        0x2e9b3a042dc7e140
    End
    RecordSource ="Competitors-Temp"
    Caption ="Competitors"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =6696
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    BackStyle =0
                    Left =8990
                    Top =1140
                    Width =780
                    Height =256
                    TabIndex =11
                    BackColor =12632256
                    BorderColor =12632256
                    Name ="PIN"
                    ControlSource ="PIN"
                    StatusBarText ="Personal ID Number"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =8050
                            Top =1133
                            Width =795
                            Height =270
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text15"
                            Caption ="System #:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1203
                    Top =810
                    Width =2550
                    Height =256
                    TabIndex =3
                    BorderColor =12632256
                    Name ="Gname"
                    ControlSource ="Gname"
                    StatusBarText ="Given Name(s)"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Top =803
                            Width =1065
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text17"
                            Caption ="Given Name:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1203
                    Top =1170
                    Width =2550
                    Height =256
                    TabIndex =4
                    BorderColor =12632256
                    Name ="Surname"
                    ControlSource ="Surname"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =51
                            Top =1170
                            Width =1008
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text19"
                            Caption ="Surname:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =2534
                    Top =1518
                    Width =1215
                    Height =256
                    TabIndex =6
                    BorderColor =12632256
                    HelpContextId =10000
                    Name ="DOB"
                    ControlSource ="DOB"
                    StatusBarText ="Date of Birth"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =1814
                            Top =1518
                            Width =576
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text25"
                            Caption ="DOB:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =144
                    Top =2692
                    Width =2805
                    Height =3795
                    TabIndex =8
                    BorderColor =12632256
                    Name ="Comments"
                    ControlSource ="Comments"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =2
                            Left =144
                            Top =2376
                            Width =1008
                            Height =240
                            FontWeight =400
                            BackColor =8421631
                            Name ="Text29"
                            Caption ="Comments:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5328
                    Top =90
                    Width =1620
                    TabIndex =12
                    BorderColor =12632256
                    Name ="State"
                    ControlSource ="ID"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =4335
                            Top =94
                            Width =855
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text37"
                            Caption ="ID Number:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =1197
                    Top =435
                    Width =1035
                    TabIndex =1
                    BorderColor =12632256
                    Name ="Postcode"
                    ControlSource ="Postcode"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =390
                            Top =435
                            Width =663
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text39"
                            Caption ="ID #"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5328
                    Top =1541
                    Width =2925
                    TabIndex =13
                    BorderColor =12632256
                    Name ="Hphone"
                    ControlSource ="Hphone"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =85
                            Left =3855
                            Top =1541
                            Width =1335
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text41"
                            Caption ="Home Phone:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5328
                    Top =1901
                    Width =2925
                    TabIndex =14
                    BorderColor =12632256
                    Name ="Wphone"
                    ControlSource ="Wphone"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =85
                            Left =3885
                            Top =1901
                            Width =1305
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text43"
                            Caption ="Work Phone:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =4150
                    Left =1194
                    Top =72
                    Width =2560
                    Height =256
                    BorderColor =12632256
                    ColumnInfo ="\"\";\">\";\"\";\"\";\"10\";\"14\""
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT House.H_Code, House.H_NAme, House.Include FROM House WHERE ((House.Includ"
                        "e=Yes)) ORDER BY House.H_Code;"
                    ColumnWidths ="885;3015"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =165
                            Top =74
                            Width =870
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text46"
                            Caption ="Team:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =3096
                    Top =2664
                    Width =5390
                    Height =3779
                    TabIndex =10
                    Name ="SFevents"
                    SourceObject ="Form.Competitors Subform"
                    LinkChildFields ="PIN"
                    LinkMasterFields ="PIN"

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3850
                    Top =435
                    TabIndex =2
                    BorderColor =12632256
                    Name ="Field60"
                    ControlSource ="Include"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="Yes"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =87
                            Left =2236
                            Top =435
                            Width =1515
                            Height =256
                            FontWeight =400
                            Name ="Text61"
                            Caption ="Include in Carnival:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =1150
                    Left =1201
                    Top =1546
                    Width =550
                    Height =256
                    TabIndex =5
                    BorderColor =12632256
                    Name ="Sex"
                    ControlSource ="Sex"
                    RowSourceType ="Value List"
                    RowSource ="\"M\";\"Male\";\"F\";\"Female\";\"-\";\"Both\""
                    ColumnWidths ="270;630"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =312
                            Top =1546
                            Width =765
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text67"
                            Caption ="Sex"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8640
                    Top =5112
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =9
                    ForeColor =8404992
                    Name ="SaveBut"
                    Caption ="Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    ListWidth =955
                    Left =1190
                    Top =1886
                    Width =955
                    Height =256
                    TabIndex =7
                    BorderColor =12632256
                    HelpContextId =10000
                    ColumnInfo ="\"\";\"\";\"2\";\"1\""
                    Name ="Age"
                    ControlSource ="Age"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Competitors.Age FROM Competitors ORDER BY Competitors.Age;"
                    ColumnWidths ="705"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =511
                            Top =1886
                            Width =525
                            Height =256
                            FontWeight =400
                            Name ="Text69"
                            Caption ="Age:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8640
                    Top =144
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =15
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
                    Left =8640
                    Top =5976
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =16
                    ForeColor =128
                    Name ="CancelBut"
                    Caption ="Cancel"
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
                    TextAlign =1
                    Left =3456
                    Top =2376
                    Width =960
                    Height =210
                    FontWeight =400
                    BackColor =8421631
                    Name ="Text50"
                    Caption ="Event"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =5472
                    Top =2376
                    Width =435
                    Height =225
                    FontWeight =400
                    BackColor =8421631
                    Name ="Text51"
                    Caption ="Final"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =6336
                    Top =2376
                    Width =360
                    Height =225
                    FontWeight =400
                    BackColor =8421631
                    Name ="Text52"
                    Caption ="Pl."
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =7632
                    Top =2376
                    Width =570
                    Height =225
                    FontWeight =400
                    BackColor =8421631
                    Name ="Text53"
                    Caption ="Pts"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    TextAlign =2
                    Left =6840
                    Top =2376
                    Width =795
                    Height =225
                    FontWeight =400
                    BackColor =8421631
                    Name ="Text64"
                    Caption ="Result"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =5904
                    Top =2376
                    Width =435
                    Height =225
                    FontWeight =400
                    BackColor =8421631
                    Name ="Text71"
                    Caption ="Heat"
                    FontName ="Arial"
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

Private Sub Address1_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Address2_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Age_AfterUpdate()

    GlobalChange = True

  If IsNull(Me![DOB]) Then
    If IsNumeric(Forms![Competitors]![Age]) Then
        Forms![Competitors]![DOB] = "1/1/" & Year(Now) - Forms![Competitors]![Age]
    Else
        Forms![Competitors]![DOB] = "1/1/01"
    End If
  End If

 
End Sub

Private Sub Age_BeforeUpdate(Cancel As Integer)

    
    Cancel = False
    Dim MyDb As Database

    OriginalValue = Forms![Competitors]![Age].OldValue
    x = Forms![Competitors]![Age]

    If IsNull(x) Then
        Response = MsgBox("You must enter a value for the competitors age", vbInformation)

    ElseIf Trim(Str(Val(x))) <> Trim(x) Then
        Response = MsgBox("The age must be numeric.", vbInformation)
    Else
        If Not (IsNull(OriginalValue)) Then '**** ie not the first entry
            
            Conti = MsgBox("This competitor will be removed from all events if this value is changed.  This action cannot be undone.  Do you want to continue?", vbYesNo + vbDefaultButton2 + vbInformation)
            
            If Conti = 6 Then
                
                Set MyDb = DBEngine.Workspaces(0).Databases(0)
                q1 = "Delete * from CompEvents Where PIN="
                q1 = q1 & Me![PIN]
                MyDb.Execute (q1)
    
            Else
                Cancel = True
                'Forms![Competitors]![Age] = OriginalValue
            End If
    
        End If
    End If

End Sub

Private Sub Button55_Click()
On Error GoTo Err_Button55_Click

    If Not CheckDataValidity() Then
        DoCmd.Close
    End If
                                       
Exit_Button55_Click:
    Exit Sub

Err_Button55_Click:
    MsgBox Error$
    Resume Exit_Button55_Click
    
End Sub

Private Sub Button56_Click()
On Error GoTo Err_Button56_Click
    
    DoCmd.GoToRecord , , A_NEWREC

Exit_Button56_Click:
    Exit Sub

Err_Button56_Click:
    MsgBox Error$
    Resume Exit_Button56_Click
    
End Sub

Private Sub Button57_Click()
On Error GoTo Err_Button57_Click


    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDelete

Exit_Button57_Click:
    Exit Sub

Err_Button57_Click:
    MsgBox Error$
    Resume Exit_Button57_Click
    
End Sub

Private Sub Button65_Click()
On Error GoTo Err_Button65_Click


    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdDelete

Exit_Button65_Click:
    Exit Sub

Err_Button65_Click:
    MsgBox Error$
    Resume Exit_Button65_Click
    
End Sub

Private Sub CancelBut_Click()

    If GlobalChange Then
        Response = MsgBox("Changes have been made to this competitor.  Are you sure you want to cancel and lose these changes?", vbYesNo + vbDefaultButton2 + vbInformation, "Confirm Cancel")
        If Response = vbYes Then
            GlobalCancel = True
            DoCmd.Close
        End If
    Else
        GlobalCancel = True
        DoCmd.Close
    End If

End Sub

Private Function CheckDataValidity()

    Cancel = False
    If IsNull(Me![Age]) Then
        Response = MsgBox("You must enter an age.", vbInformation)
        Cancel = True
    ElseIf IsNull(Me![Sex]) Then
        Response = MsgBox("You must enter a sex.", vbInformation)
        Cancel = True
        
    ElseIf IsNull(Me![H_Code]) Then
        Response = MsgBox("You must enter a team.", vbInformation)
        Cancel = True
    End If

    CheckDataValidity = Cancel

End Function

Private Sub Comments_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub DOB_AfterUpdate()

    GlobalChange = True

      If IsNumeric(Me![Age]) Or IsNull(Me![Age]) Then
        y = Format$(Me![DOB], "YYYY")
        Yn = Format$(Now, "YYYY")
        If Yn > y Then
            Me![Age] = Yn - y
        End If

    End If
End Sub

Private Sub Field60_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Field66_BeforeUpdate(Cancel As Integer)
    Cancel = False
    Dim MyDb As Database

    OriginalValue = Forms![Competitors]![Age].OldValue
    x = Forms![Competitors]![Age]

    If Not (IsNull(OriginalValue)) Then '**** ie not the first entry
        
        Conti = MsgBox("This competitor will be removed from all events if this value is changed.  This action cannot be undone.  Do you want to continue?", vbYesNo + vbDefaultButton2 + vbExclamation)
        
        If Conti = vbYes Then
            
            Set MyDb = DBEngine.Workspaces(0).Databases(0)
            q1 = "Delete * from CompEvents Where PIN="
            q1 = q1 & Forms![Competitors]![PIN]
            MyDb.Execute (q1)

        Else
            Cancel = True
            Forms![Competitors]![Age] = OriginalValue
        End If

    End If

End Sub

Private Sub FindBut_Click()
On Error GoTo Err_FindBut_Click

    If Forms![Compet_Search].visible = False Then
        Forms![Compet_Search].visible = True
    Else
        Forms![Compet_Search].visible = False
    End If



Exit_FindBut_Click:
    Exit Sub

Err_FindBut_Click:
    MsgBox Error$
    Resume Exit_FindBut_Click
    
End Sub

Private Function FindFirstName(fn)

End Function


Private Sub Form_BeforeUpdate(Cancel As Integer)

    Cancel = CheckDataValidity()


End Sub

Private Sub Form_Load()

    If Me.OpenArgs = "EDIT" Then
        Me.DefaultEditing = 4
    'ElseIf Me.Openargs = "ADD" Then
    '    Me.DefaultEditing = 1

    ElseIf Me.OpenArgs = "NOT IN LIST" Then
        Me.DefaultEditing = 1
        [Age] = Forms![EnterCompetitors]![AgeFld]
        [Sex] = Forms![EnterCompetitors]![SexFld]
        [H_Code] = Forms![EnterCompetitors]![EC_Subform].Form![H_Code]
        GoSub GetNames
        [Gname] = FirstName
        [Surname] = LastName


    End If

    GoTo ExitFL
    
GetNames:

    fn = GlobalVariable
    L = Len(fn)
    LastName = ""
    FirstName = ""
    First = False

    For i = 1 To L
        Ch = Mid(fn, i, 1)
        If Ch = "," Then
            First = True
        Else
            If First Then
                FirstName = FirstName & Ch
            Else
                LastName = LastName & Ch
            End If
        End If
    Next i

    Return

ExitFL:

End Sub

Private Sub Gname_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub H_Code_AfterUpdate()

    GlobalChange = True
    
End Sub

Private Sub Hphone_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Postcode_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub SaveBut_Click()

On Error GoTo SaveBut_Click_Err
    ' Check for valid data
    If IsNull(Me![Age]) Then
        Response = MsgBox("You must enter an age.", vbInformation)
        Cancel = True
    
    ElseIf IsNull(Me![Sex]) Then
        Response = MsgBox("You must enter a sex.", vbInformation)
        Cancel = True
        
    ElseIf IsNull(Me![H_Code]) Then
        Response = MsgBox("You must enter a team.", vbInformation)
        Cancel = True
    
    ElseIf IsNull(Me![Gname]) Then
        Response = MsgBox("You must enter a first name.", vbInformation)
        Cancel = True
    
    ElseIf IsNull(Me![Surname]) Then
        Response = MsgBox("You must enter a surname.", vbInformation)
        Cancel = True
    Else
        DoCmd.Close
    End If


SaveBut_Click_Exit:
    Exit Sub

SaveBut_Click_Err:
    MsgBox (Error$)
    GoTo SaveBut_Click_Exit


End Sub

Private Sub Sex_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub State_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Suburb_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Surname_AfterUpdate()

    GlobalChange = True

End Sub

Private Sub Wphone_AfterUpdate()

    GlobalChange = True

End Sub
