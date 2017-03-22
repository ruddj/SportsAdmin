Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =9411
    ItemSuffix =71
    Left =3255
    Top =2325
    Right =12660
    Bottom =7455
    RecSrcDt = Begin
        0xff4a5ef6abcde140
    End
    RecordSource ="SELECT DISTINCTROW Events.E_Code, EventType.ET_Des, EventType.Units, Events.Sex,"
        " Events.Age, Events.nRecord, EventType.ET_Code, EventType.ET_Des, EventType.Unit"
        "s, Events.Sex, Events.Age FROM EventType INNER JOIN Events ON EventType.ET_Code "
        "= Events.ET_Code;"
    Caption ="Event Record"
    HelpFile ="SportsAdmin.chm"
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
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4136
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =5157
                    Top =283
                    Width =435
                    ColumnWidth =825
                    FontWeight =700
                    BackColor =-2147483633
                    Name ="Sex"
                    ControlSource ="Sex"
                    StatusBarText ="Male (M) / Female (F)"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4592
                            Top =283
                            Width =452
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text19"
                            Caption ="Sex:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =6374
                    Top =283
                    Width =870
                    ColumnWidth =1020
                    FontWeight =700
                    TabIndex =1
                    BackColor =-2147483633
                    Name ="Age"
                    ControlSource ="Age"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5696
                            Top =283
                            Width =565
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text21"
                            Caption ="Age:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    Left =1041
                    Top =276
                    Width =3390
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483633
                    Name ="ET_Des"
                    ControlSource ="ET_Des"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Arial"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =277
                            Top =281
                            Width =645
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text36"
                            Caption ="Event:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =255
                    Width =7755
                    Height =4136
                    Name ="Box40"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7937
                    Top =3458
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    Name ="Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =247
                    Left =173
                    Top =113
                    Width =7431
                    Height =615
                    Name ="Box45"
                End
                Begin Subform
                    OverlapFlags =247
                    SpecialEffect =3
                    Left =360
                    Top =1363
                    Width =7125
                    Height =2443
                    TabIndex =4
                    Name ="EventRecordsSF2"
                    SourceObject ="Form.EventRecordsSF"
                    LinkChildFields ="Records.E_Code"
                    LinkMasterFields ="E_Code"

                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =417
                    Top =1127
                    Width =821
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text62"
                    Caption ="Surname"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =1608
                    Top =1127
                    Width =975
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text63"
                    Caption ="Given Name"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =2891
                    Top =1125
                    Width =1125
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text64"
                    Caption ="Team"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =4726
                    Top =1127
                    Width =645
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text65"
                    Caption ="Result"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =5973
                    Top =1127
                    Width =540
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text66"
                    Caption ="Date"
                    FontName ="Arial"
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =255
                    Left =226
                    Top =1072
                    Width =7371
                    Height =2910
                    Name ="Box67"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =285
                    Top =793
                    Width =1605
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text68"
                    Caption ="Event Record History"
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

Option Explicit

Private Sub Add_Click()

    Dim Continue As Variant, Response As Variant

    Dim Criteria As String, rs As Recordset
    Dim nValu As String, res As String, nUnit As String
    Dim success As Boolean
    
    Set rs = CurrentDb.OpenRecordset("Records", dbOpenDynaset)   ' Create dynaset.
    
    If IsNull(Me![Surname]) Then
        MsgBox ("You must enter a surname.")
    ElseIf Trim(Me![Surname]) = "" Then
        MsgBox ("You must enter a surname.")
    ElseIf IsNull(Me![Date]) Then
        MsgBox ("You must enter a date.")
    ElseIf Trim(Me![Date]) = "" Then
        MsgBox ("You must enter a date.")
    ElseIf IsNull(Me![Record]) Then
        MsgBox ("You must enter a result.")
    ElseIf Trim(Me![Record]) = "" Then
        MsgBox ("You must enter a result.")
    Else
        Continue = True
        res = Me![Record]
        nValu = ""
        nUnit = Me![nUnit]
        Call Calculate_Results(res, nValu, nUnit, success)
        If Not (Better(Val(res), Me![E_Code])) Then
            Response = MsgBox("The event record you are about to add is not better than the existing record.  Do you want to continue?", 20, "Record Integrity Violation")
            If Response <> 6 Then Continue = False
        End If
        
        If Me![Date] < DMax("[Date]", "Records", "[E_Code]=" & Me![E_Code]) Then
            Response = MsgBox("The date for the event record you are about to add is older than the most recent record.  Do you want to continue?", 20, "Record Integrity Violation")
            If Response <> 6 Then Continue = False
        End If

        If Continue Then
            Criteria = "[E_Code]=" & Me![E_Code] & " AND [Date]=#" & Format(Me![Date], "mm/dd/yy") & "#"
        
            rs.FindFirst Criteria    ' Find first occurrence.
            If rs.NoMatch Then
                rs.AddNew
                rs!E_Code = Me![E_Code]
                rs!Surname = Me![Surname]
                rs!Gname = Me![Gname]
                rs!H_Code = Me![House]
                rs!Date = Me![Date]
                rs!Result = Me![Record]
                rs!nResult = Val(res)
                rs.Update
                
            Else
    
                Response = MsgBox("Do you want to replace the existing record?", 20, "Replace Record")
                If Response = 6 Then
                    rs.Edit
                    rs!E_Code = Me![E_Code]
                    rs!Surname = Me![Surname]
                    rs!Gname = Me![Gname]
                    rs!H_Code = Me![House]
                    rs!Date = Me![Date]
                    rs!Result = Me![Record]
                    rs!nResult = Val(res)
                    rs.Update

                End If
                
            End If
                
                Me![EventRecordsSF].Requery
                Me![EventRecordsSF2].Requery
            End If
        
    End If

    rs.Close
End Sub

Private Sub Close_Click()
On Error GoTo Err_Close_Click

    Call SetCurrentRecord

    DoCmd.Close

Exit_Close_Click:
    Exit Sub

Err_Close_Click:
    MsgBox Error$
    Resume Exit_Close_Click
    
End Sub

Private Sub PIN_AfterUpdate()

    Dim Criteria As Variant

    If Not IsNull(Me![PIN]) Then
        Criteria = "[PIN]=" & Me![PIN]
    
        Me![Gname] = DLookup("[Gname]", "CompetitorsFull", Criteria)
        Me![Surname] = DLookup("[Surname]", "CompetitorsFull", Criteria)
        Me![House] = DLookup("[H_Code]", "CompetitorsFull", Criteria)
    End If
    
End Sub

Private Sub Record_AfterUpdate()

    'Stop

  Dim res As String
  Dim Runit As String
  Dim nValu As String
  Dim success As Boolean
  
  res = Me![Record]
  
  If Not (IsNull(res)) Then
    
    nValu = ""
    Runit = Forms![EventType]![Units]
    Call Calculate_Results(res, nValu, Runit, success)

    Me![Record] = nValu
    Me![nRecord] = res

  Else
    ' When the Result (time or distance or points) is set to NULL then
    ' set Numeric Result and Place to 0

    Me.[nRecord] = 0
    
  End If
    
End Sub

Private Sub SetCurrentRecord()

    Dim RecDate As Variant, Criteria As Variant, res As Variant, nRes As Variant, Gname As Variant


    RecDate = DLookup("[MaxOfDate]", "Record-Most Recent", "[E_Code]=" & Me![E_Code])
'    If Not IsNull(RecDate) Then
'        Criteria = "[E_Code]=" & Me![E_Code] & " AND [Date]=#" & Format(RecDate, "mm/dd/yy") & "#"
'        Res = DLookup("[Result]", "Records", Criteria)
'        nRes = DLookup("[nResult]", "Records", Criteria)
'        Gname = DLookup("[Gname]", "Records", Criteria)
'        Sname = DLookup("[Surname]", "Records", Criteria)
'        RecName = UCase(Sname) & ", " & Gname
'        Hcode = DLookup("[H_code]", "Records", Criteria)
'        Hid = DLookup("[H_id]", "House", "[H_Code]=""" & Hcode & """")
        
        'Forms![EventType]![ET_Sub1].Form![Record] = Res
        'Forms![EventType]![ET_Sub1].Form![nRecord] = nRes

        'q = "UPDATE DISTINCTROW Events SET "
        'q = q & " Events.nRecord=" & nRes & ", Events.Record=""" & Res & """, Events.RecName=""" & RecName & """, Events.RecHouse=" & Hid
        'q = q & " where Events.E_Code=" & Me![E_Code]
        
        'DoCmd SetWarnings False
        'DoCmd RunSQL q
        'DoCmd SetWarnings True
'    End If
    
End Sub
