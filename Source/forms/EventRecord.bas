Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =9240
    ItemSuffix =73
    Left =7935
    Top =2640
    Right =17865
    Bottom =11760
    HelpContextId =280
    Filter ="[E_Code]=2"
    RecSrcDt = Begin
        0xd19592b911cde140
    End
    RecordSource ="SELECT DISTINCTROW Events.E_Code, EventType.ET_Des, EventType.Units, Events.Sex,"
        " Events.Age, Events.nRecord, EventType.ET_Code, EventType.ET_Des, EventType.Unit"
        "s, Events.Sex, Events.Age FROM EventType INNER JOIN Events ON EventType.ET_Code "
        "= Events.ET_Code;"
    Caption ="Event Record"
    HelpFile ="SportsAdmin.chm"
    FilterOnLoad =255
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
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
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
            Height =6768
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    Left =5157
                    Top =283
                    Width =435
                    ColumnWidth =825
                    FontWeight =700
                    TabIndex =6
                    BackColor =12632256
                    Name ="Sex"
                    ControlSource ="Sex"
                    StatusBarText ="Male (M) / Female (F)"
                    FontName ="Tahoma"

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
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =3
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    Left =6374
                    Top =283
                    Width =870
                    ColumnWidth =1020
                    FontWeight =700
                    TabIndex =7
                    BackColor =12632256
                    Name ="Age"
                    ControlSource ="Age"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

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
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =3
                    Left =1539
                    Top =2267
                    Width =870
                    Height =225
                    TabIndex =3
                    BorderColor =12632256
                    Name ="Record"
                    StatusBarText ="Record for this event stored as a text value"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =635
                            Top =2267
                            Width =791
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text27"
                            Caption ="Record:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1550
                    Top =1247
                    Width =3075
                    Height =225
                    BorderColor =12632256
                    Name ="Gname"
                    StatusBarText ="name of Competitor who holds current record"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =257
                            Top =1252
                            Width =1185
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text29"
                            Caption ="Given Name:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =3130
                    Left =1542
                    Top =1927
                    Width =3070
                    Height =227
                    TabIndex =2
                    BorderColor =12632256
                    ColumnInfo ="\"\";\">\";\"\";\"\";\"10\";\"100\""
                    Name ="House"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT House.H_Code, House.H_NAme FROM House ORDER BY House.H_NAme;"
                    ColumnWidths ="0;1441"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =541
                            Top =1927
                            Width =885
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text33"
                            Caption ="Team:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =1041
                    Top =276
                    Width =3390
                    FontWeight =700
                    TabIndex =8
                    BackColor =12632256
                    Name ="ET_Des"
                    ControlSource ="ET_Des"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

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
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    Left =2412
                    Top =2267
                    Width =510
                    Height =225
                    TabIndex =9
                    BackColor =-2147483633
                    Name ="nUnit"
                    ControlSource ="Units"
                    StatusBarText ="Record for this event stored as a text value"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8000
                    Top =6082
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =10
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Subform
                    Enabled = NotDefault
                    OverlapFlags =215
                    OldBorderStyle =0
                    Left =348
                    Top =3404
                    Width =7095
                    Height =343
                    TabIndex =11
                    Name ="EventRecordsSF"
                    SourceObject ="Form.EventRecordsSF"
                    LinkChildFields ="Records.E_Code"
                    LinkMasterFields ="E_Code"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =6402
                    Top =1247
                    Width =855
                    TabIndex =12
                    BorderColor =12632256
                    Name ="E_Code"
                    ControlSource ="E_Code"
                    StatusBarText ="Record for this event"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5385
                            Top =1247
                            Width =904
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text44"
                            Caption ="E_Code"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =247
                    Left =173
                    Top =113
                    Width =7386
                    Height =615
                    Name ="Box45"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1550
                    Top =1587
                    Width =3075
                    Height =225
                    TabIndex =1
                    BorderColor =12632256
                    Name ="Surname"
                    StatusBarText ="name of Competitor who holds current record"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =1585
                            Width =1215
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text47"
                            Caption ="Surname:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    TextFontFamily =34
                    Left =4875
                    Top =2097
                    Width =2244
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    Name ="Add"
                    Caption ="&Add New Record"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =3624
                    Top =2269
                    Width =1005
                    Height =225
                    TabIndex =4
                    BorderColor =12632256
                    Name ="Date"
                    Format ="Short Date"
                    StatusBarText ="name of Competitor who holds current record"
                    DefaultValue ="=Date()"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3035
                            Top =2267
                            Width =480
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text50"
                            Caption ="Date:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =736
                    Top =3168
                    Width =791
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text51"
                    Caption ="Surname"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1927
                    Top =3168
                    Width =945
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text52"
                    Caption ="Given Name"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3210
                    Top =3165
                    Width =1095
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text54"
                    Caption ="Team"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5045
                    Top =3168
                    Width =615
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text55"
                    Caption ="Result"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =6292
                    Top =3168
                    Width =510
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text56"
                    Caption ="Date"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =226
                    Top =3113
                    Width =7386
                    Height =735
                    Name ="Box57"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =240
                    Top =2850
                    Width =1605
                    Height =240
                    BackColor =-2147483633
                    Name ="Text58"
                    Caption ="Current Record"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    ColumnCount =3
                    ListWidth =2065
                    Left =1526
                    Top =850
                    Width =3130
                    Height =223
                    TabIndex =13
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="PIN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Competitors.PIN, UCase([Surname]) & \", \" & [Gname] AS Name, House.Inclu"
                        "de FROM House INNER JOIN Competitors ON House.H_Code = Competitors.H_Code WHERE "
                        "((House.Include=Yes)) ORDER BY UCase([Surname]) & \", \" & [Gname];"
                    ColumnWidths ="0;1815;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =850
                            Width =1260
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text60"
                            Caption ="Add Competitor:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =348
                    Top =4595
                    Width =7095
                    Height =1813
                    TabIndex =14
                    BorderColor =12632256
                    Name ="EventRecordsSF2"
                    SourceObject ="Form.EventRecordsSF"
                    LinkChildFields ="Records.E_Code"
                    LinkMasterFields ="E_Code"

                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =736
                    Top =4359
                    Width =791
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text62"
                    Caption ="Surname"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1927
                    Top =4359
                    Width =945
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text63"
                    Caption ="Given Name"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3210
                    Top =4365
                    Width =1095
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text64"
                    Caption ="Team"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5045
                    Top =4359
                    Width =615
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text65"
                    Caption ="Result"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =6292
                    Top =4359
                    Width =510
                    Height =225
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text66"
                    Caption ="Date"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =226
                    Top =4304
                    Width =7386
                    Height =2280
                    Name ="Box67"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =285
                    Top =4025
                    Width =2265
                    Height =240
                    BackColor =-2147483633
                    Name ="Text68"
                    Caption ="Record History"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    Left =1527
                    Top =2551
                    Width =870
                    Height =225
                    TabIndex =15
                    BorderColor =12632256
                    Name ="nRecord"
                    StatusBarText ="Record for this event stored as a text value"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =623
                            Top =2551
                            Width =791
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text70"
                            Caption ="Record:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    Left =7755
                    Width =0
                    Height =6768
                    Name ="Line71"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =72
                    TextFontFamily =34
                    Left =7983
                    Top =4965
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =16
                    HelpContextId =280
                    Name ="HelpButton"
                    Caption ="&Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

    Dim Criteria As String, db As Database, Rs As Recordset
    Dim nValu As String, res As String, nUnit As String
    Dim success As Boolean
    
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set Rs = db.OpenRecordset("Records", dbOpenDynaset)   ' Create dynaset.
    
    If IsNull(Me![Surname]) Then
        Response = MsgBox("You must enter a surname.", vbInformation)
    ElseIf Trim(Me![Surname]) = "" Then
        Response = MsgBox("You must enter a surname.", vbInformation)
    ElseIf IsNull(Me![Date]) Then
        Response = MsgBox("You must enter a date.", vbInformation)
    ElseIf Trim(Me![Date]) = "" Then
        Response = MsgBox("You must enter a date.", vbInformation)
    ElseIf IsNull(Me![Record]) Then
        Response = MsgBox("You must enter a result.", vbInformation)
    ElseIf Trim(Me![Record]) = "" Then
        Response = MsgBox("You must enter a result.", vbInformation)
    Else
        Continue = True
        res = Me![Record]
        nValu = ""
        nUnit = Me![nUnit]
        Call Calculate_Results(res, nValu, nUnit, success)
        If Not (Better(Val(res), Me![E_Code])) Then
            Response = MsgBox("The event record you are about to add is not better than the existing record.  Do you want to continue?", vbYesNo + vbCritical, "Record Integrity Violation")
            If Response <> vbYes Then Continue = False
        End If
        
        If Me![Date] < DMax("[Date]", "Records", "[E_Code]=" & Me![E_Code]) Then
            Response = MsgBox("The date for the event record you are about to add is older than the most recent record.  Do you want to continue?", vbYesNo + vbCritical, "Record Integrity Violation")
            If Response <> vbYes Then Continue = False
        End If

        If Continue Then
            Criteria = "[E_Code]=" & Me![E_Code] & " AND [Date]=#" & Format(Me![Date], "mm/dd/yy") & "#"
        
            Rs.FindFirst Criteria    ' Find first occurrence.
            If Rs.NoMatch Then
                Rs.AddNew
                Rs!E_Code = Me![E_Code]
                Rs!Surname = Me![Surname]
                Rs!Gname = Me![Gname]
                Rs!H_Code = Me![House]
                Rs!Date = Me![Date]
                Rs!Result = Me![Record]
                Rs!nResult = Val(res)
                Rs.Update
                
            Else
    
                Response = MsgBox("A record has already been set on this day.  Do you want to replace the existing record?", vbYesNo + vbInformation, "Replace Record")
                If Response = vbYes Then
                    Rs.Edit
                    Rs!E_Code = Me![E_Code]
                    Rs!Surname = Me![Surname]
                    Rs!Gname = Me![Gname]
                    Rs!H_Code = Me![House]
                    Rs!Date = Me![Date]
                    Rs!Result = Me![Record]
                    Rs!nResult = Val(res)
                    Rs.Update

                End If
                
            End If
                
                Me![EventRecordsSF].Requery
                Me![EventRecordsSF2].Requery
            End If
        
    End If

    Rs.Close
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

On Error GoTo Record_AfterUpdate_Err
    'Stop

  Dim res As String, ET_Code As Long
  Dim Runit As String
  Dim nValu As String
  Dim success As Boolean
  
  res = Me![Record]
  
  If Not (IsNull(res)) Then
    
    nValu = ""
    
    ET_Code = DLookup("[ET_Code]", "Events", "[E_Code]=" & Me!E_Code)
    If IsNull(ET_Code) Then
      MsgBox ("An unexpected error has occured in [Record_AfterUpdate]: ET_Code is null")
    Else
      
      Runit = DLookup("[Units]", "EventType", "[ET_Code]=" & ET_Code) 'Forms![EventType]![Units]
      If IsNull(ET_Code) Then
        MsgBox ("An unexpected error has occured in [Record_AfterUpdate]: Units is null")
      Else
        Call Calculate_Results(res, nValu, Runit, success)
        Me![Record] = nValu
        Me![nRecord] = res
      End If
    End If
  Else
    ' When the Result (time or distance or points) is set to NULL then
    ' set Numeric Result and Place to 0

    Me.[nRecord] = 0
    
  End If
    
Record_AfterUpdate_Exit:
  Exit Sub
  
Record_AfterUpdate_Err:
  MsgBox ("An error has occured in [Record_AfterUpdate]: " & Err.Description)
  GoTo Record_AfterUpdate_Exit
  
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
