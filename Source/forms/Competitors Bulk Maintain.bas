Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =20
    Width =10152
    ItemSuffix =43
    Left =7470
    Top =2460
    Right =18315
    Bottom =10890
    HelpContextId =70
    RecSrcDt = Begin
        0x2231f2280fcde140
    End
    RecordSource ="MiscellaneousLocal"
    Caption ="Bulk Maintenance of Competitors"
    HelpFile ="SportsAdmin.chm"
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7325
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8784
                    Top =6494
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Close"
                    Caption ="Close"
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
                    Left =8856
                    Top =216
                    Width =1119
                    Height =675
                    FontSize =7
                    FontWeight =400
                    TabIndex =2
                    Name ="Delete"
                    Caption ="Delete Tagged Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =4362
                    Top =963
                    Width =660
                    Height =210
                    Name ="Text14"
                    Caption ="House"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =5317
                    Top =963
                    Width =570
                    Height =225
                    Name ="Text15"
                    Caption ="Sex"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =6119
                    Top =963
                    Width =450
                    Height =225
                    Name ="Text16"
                    Caption ="Age"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =793
                    Top =963
                    Width =1155
                    Height =225
                    Name ="Text20"
                    Caption ="Given Name"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =2608
                    Top =963
                    Width =1230
                    Height =225
                    Name ="Text21"
                    Caption ="Surname"
                    FontName ="Tahoma"
                End
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =3
                    Left =453
                    Top =1206
                    Width =8053
                    Height =5956
                    Name ="Compet"
                    SourceObject ="Form.Competitors Bulk Maintain SF"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =6856
                    Top =963
                    Width =555
                    Height =225
                    Name ="Text30"
                    Caption ="DOB"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8784
                    Top =5814
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    HelpContextId =70
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =85
                    ColumnCount =2
                    ListWidth =1418
                    Left =1700
                    Top =282
                    Width =1690
                    TabIndex =4
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Field1"
                    ControlSource ="CompBulkField"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCTROW [Table Field Names].FieldName, [Table Field Names].DisplayNam"
                        "e, [Table Field Names].Table FROM [Table Field Names] WHERE (([Table Field Names"
                        "].Table=\"Competitors\")) ORDER BY [Table Field Names].DisplayName;"
                    ColumnWidths ="0;1441"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =963
                            Top =282
                            Width =630
                            Height =240
                            Name ="Text35"
                            Caption ="Field:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =2
                    ListWidth =1418
                    Left =3514
                    Top =282
                    Width =955
                    Height =224
                    TabIndex =5
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="Op1"
                    ControlSource ="CompBulkOp"
                    RowSourceType ="Table/Query"
                    RowSource ="Select [Operator] From [Operators];"
                    ColumnWidths ="1440"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =85
                    Left =4592
                    Top =282
                    TabIndex =6
                    BorderColor =12632256
                    Name ="Value1"
                    ControlSource ="CompBulkValue"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7200
                    Top =216
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    Name ="Update Display"
                    Caption ="Update Display"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    EventProcPrefix ="Update_Display"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    Left =7653
                    Top =963
                    Width =555
                    Height =225
                    Name ="Text42"
                    Caption ="Tag"
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
Option Explicit

Private Sub Button27_Click()
On Error GoTo Err_Button27_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Button27_Click:
    Exit Sub

Err_Button27_Click:
    MsgBox Error$
    Resume Exit_Button27_Click
    
End Sub

Private Sub Close_Click()
On Error GoTo Err_Close_Click


    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

Private Function CompetitorEnrolled(Gname, Sname)

    If IsNull(Gname) And IsNull(Sname) Then
        CompetitorEnrolled = False
    Else
        CompetitorEnrolled = True
    End If

End Function

Private Sub Delete_Click()

    Dim Response As Variant, W As Variant, Q As Variant, DelComp As Variant, W2 As Variant

    On Error GoTo Err_Delete_Click

    W = WhereClause()
    If W = "" Then
        W = " WHERE [Include]=Yes "
    Else
        W = W & " AND [Include]=Yes"
    End If
        

    'Stop
    'W2 = Left(W, InStr(W, "competitors.") - 1) & Right(W, Len(W) - (InStr(W, "competitors.") + Len("competitors.") - 1))
    W2 = Right(W, Len(W) - 6)

    DelComp = DCount("[PIN]", "Competitors", W2)
    
    Response = MsgBox("Are you sure you wish to delete " & Str(DelComp) & " competitors?", vbYesNo + vbInformation + vbDefaultButton2, "Delete Competitors")
    If Response = vbYes Then
        Q = "DELETE DISTINCTROW Competitors.Gname, Competitors.Surname, Competitors.Sex, Competitors.H_Code, Competitors.DOB, Competitors.Age "
        Q = Q & "FROM Competitors "

        Q = Q & W
        'Stop
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True
        Me![Compet].Requery

    End If

Exit_Delete_Click:
    DoCmd.SetWarnings True
    Exit Sub

Err_Delete_Click:
    MsgBox (Error)
    GoTo Exit_Delete_Click

End Sub

Private Sub Form_Load()

    Update_Display_Click

End Sub

Private Sub ImportData_Click()
    
End Sub

Private Sub Locate_Click()


    'Dim n As Variant

    'n = GetFileName("Competitor Text File", "Text Files (*.txt)|*.txt|All (*.*)|*.*||", 1, "*.txt")
    'If n <> NoFileSelection Then
    '    Me![FullFileName] = Trim(n)
    'End If

End Sub

Private Sub SortByName_AfterUpdate()

    Dim Q As Variant

    If Me![SortByName] Then
        Q = "SELECT DISTINCTROW ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age, ImportData.Sex, ImportData.HE_Code, ImportData.ET_Des, ImportData.Heat, ImportData.Competitor, ImportData.Memo "
        Q = Q & "FROM ImportData ORDER BY ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age"
        Me![I_Data].Form.RecordSource = Q
    Else
        Q = "SELECT DISTINCTROW ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age, ImportData.Sex, ImportData.HE_Code, ImportData.ET_Des, ImportData.Heat, ImportData.Competitor, ImportData.Memo "
        Q = Q & "FROM ImportData " ' ORDER BY ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age"
        Me![I_Data].Form.RecordSource = Q
    End If
    
    Me![I_Data].Form.Requery


End Sub

Private Sub Update_Display_Click()
On Error GoTo Err_Update_Display_Click

    Dim Q As Variant, q1 As Variant

    'Stop
    Q = WhereClause()
    
    q1 = "SELECT DISTINCTROW Competitors.Surname, Competitors.Gname, Competitors.H_Code, Competitors.Age, Competitors.DOB, Competitors.Sex, Competitors.Include "
    q1 = q1 & "FROM [Import Competitors], Competitors "
    q1 = q1 & Q
    q1 = q1 & " ORDER BY [Surname], Competitors.Gname, Competitors.[H_Code], Competitors.[Age], Competitors.[DOB], Competitors.[Sex] "

    Me![Compet].Form.RecordSource = q1
    Me![Compet].Requery

Exit_Update_Display_Click:
    Exit Sub

Err_Update_Display_Click:
    MsgBox Error$
    GoTo Exit_Update_Display_Click
End Sub

Private Function WhereClause()

    Dim Fld1 As Variant, Op1 As Variant, Value1 As Variant, Q As Variant, Ty As Variant, D As Variant
    Fld1 = Me![Field1]
    Op1 = Me![Op1]
    Value1 = Me![Value1]
    If Not (IsNull(Fld1)) And Not IsNull(Op1) And Not IsNull(Value1) Then
        Ty = DLookup("[Type]", "Table Field Names", "[FieldName]=""" & Fld1 & """" & " and [Table]=""Competitors""")
        Q = "competitors.[" & Fld1 & "] " & Op1
        
        
        If Ty = "Date" Then
            If IsDate(Value1) Then
                D = CVDate(Value1)
                Q = Q & " #" & Format(D, "mm/dd/yy") & "#"
            Else
                Q = Q & " *"
            End If
            
        ElseIf Ty = "Number" Then
          If IsNumeric(Value1) Then
            Q = Q & " " & Value1
          Else
            Q = Q & " """ & Value1 & """"
          
          End If
        Else
            Q = Q & " """ & Value1 & """"
        End If
        
        WhereClause = " WHERE " & Q

    Else
        WhereClause = ""
    End If
        
End Function
