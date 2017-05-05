Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
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
    BorderStyle =3
    GridY =10
    Width =8503
    ItemSuffix =144
    Left =5625
    Top =2265
    Right =14805
    Bottom =9090
    HelpContextId =260
    RecSrcDt = Begin
        0x03b3b85379f4e140
    End
    RecordSource ="Misc-EventLists"
    Caption ="Generate Event Lists"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    OnActivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    RibbonName ="SportsMenu"
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
            Width =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial"
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
            Height =5782
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =144
                    Top =144
                    Width =2408
                    Height =5511
                    Name ="EventSF"
                    SourceObject ="Form.Report SF2"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2820
                    Top =5160
                    Width =1134
                    Height =510
                    TabIndex =1
                    HelpContextId =260
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =7260
                    Top =5160
                    Width =1134
                    Height =510
                    FontWeight =700
                    TabIndex =2
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =2775
                    Top =105
                    Width =5625
                    Height =4845
                    TabIndex =3
                    Name ="TabCtl125"
                    FontName ="Tahoma"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =2910
                            Top =510
                            Width =5359
                            Height =4310
                            Name ="Page126"
                            Caption ="Common Lists"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin ComboBox
                                    RowSourceTypeInt =1
                                    OverlapFlags =223
                                    ColumnCount =2
                                    ListWidth =1375
                                    Left =4182
                                    Top =1420
                                    Width =1525
                                    Height =227
                                    BackColor =16777215
                                    BorderColor =12632256
                                    Name ="Sex_DD"
                                    ControlSource ="Rsex"
                                    RowSourceType ="Value List"
                                    RowSource ="\"*\";\"Any\";\"M\";\"Male\";\"F\";\"Female\""
                                    ColumnWidths ="391;735"
                                    OnDblClick ="[Event Procedure]"
                                    DefaultValue ="\"*\""
                                    FontName ="Tahoma"
                                    ControlTipText ="Enter a '*' for any gender."

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =3210
                                            Top =1424
                                            Width =765
                                            Height =240
                                            FontWeight =400
                                            Name ="Sex_DD_Tit"
                                            Caption ="Gender"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =223
                                    ListWidth =1510
                                    Left =4179
                                    Top =1761
                                    Width =1540
                                    Height =225
                                    TabIndex =1
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
                                    FontName ="Tahoma"
                                    ControlTipText ="Enter a '*' for all final-levels."

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =3328
                                            Top =1761
                                            Width =630
                                            Height =240
                                            FontWeight =400
                                            Name ="Text71"
                                            Caption ="Final"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =223
                                    ListWidth =1510
                                    Left =4189
                                    Top =2463
                                    Width =1510
                                    Height =225
                                    TabIndex =2
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
                                    FontName ="Tahoma"
                                    ControlTipText ="Enter a '*' for all heats,"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =3323
                                            Top =2463
                                            Width =630
                                            Height =240
                                            FontWeight =400
                                            Name ="Text73"
                                            Caption ="Heat"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =223
                                    ListWidth =1510
                                    Left =4182
                                    Top =1080
                                    Width =1525
                                    Height =225
                                    TabIndex =3
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
                                    FontName ="Tahoma"
                                    ControlTipText ="Enter a '*' for all ages."

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =3331
                                            Top =1080
                                            Width =630
                                            Height =240
                                            FontWeight =400
                                            Name ="Text75"
                                            Caption ="Age"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =4209
                                    Top =2916
                                    TabIndex =4
                                    BorderColor =12632256
                                    Name ="Future"
                                    ControlSource ="Rfuture"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =3493
                                            Top =2859
                                            Width =615
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text81"
                                            Caption ="Future"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =4211
                                    Top =3222
                                    TabIndex =5
                                    BorderColor =12632256
                                    Name ="Active"
                                    ControlSource ="Ractive"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =3495
                                            Top =3165
                                            Width =615
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text83"
                                            Caption ="Active"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =5515
                                    Top =2916
                                    TabIndex =6
                                    BorderColor =12632256
                                    Name ="Completed"
                                    ControlSource ="Rcompleted"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =4514
                                            Top =2859
                                            Width =900
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text85"
                                            Caption ="Completed"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =5517
                                    Top =3222
                                    TabIndex =7
                                    BorderColor =12632256
                                    Name ="Promoted"
                                    ControlSource ="Rpromoted"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =4516
                                            Top =3165
                                            Width =900
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text87"
                                            Caption ="Promoted"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    BackStyle =0
                                    Left =4399
                                    Top =2859
                                    Width =216
                                    TabIndex =8
                                    BackColor =-2147483633
                                    Name ="FutureEB"
                                    FontName ="Tahoma"

                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    BackStyle =0
                                    Left =4396
                                    Top =3299
                                    Width =216
                                    TabIndex =9
                                    BackColor =-2147483633
                                    Name ="ActiveEB"
                                    FontName ="Tahoma"

                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    BackStyle =0
                                    Left =5818
                                    Top =2859
                                    Width =216
                                    TabIndex =10
                                    BackColor =-2147483633
                                    Name ="CompletedEB"
                                    FontName ="Tahoma"

                                End
                                Begin TextBox
                                    Visible = NotDefault
                                    OverlapFlags =223
                                    Left =5818
                                    Top =3142
                                    Width =216
                                    TabIndex =11
                                    Name ="PromotedEB"
                                    FontName ="Tahoma"

                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =255
                                    Left =3210
                                    Top =2292
                                    Width =2778
                                    Height =1248
                                    Name ="Box93"
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =7907
                                    Top =1995
                                    TabIndex =12
                                    BorderColor =12632256
                                    Name ="Detailed"
                                    ControlSource ="Rdetailed"
                                    DefaultValue ="False"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =6540
                                            Top =1995
                                            Width =1245
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text101"
                                            Caption ="Detailed Lists"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =7907
                                    Top =2279
                                    TabIndex =13
                                    BorderColor =12632256
                                    Name ="SummaryReport"
                                    ControlSource ="Rsummary"
                                    DefaultValue ="False"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            Left =6540
                                            Top =2278
                                            Width =1245
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text115"
                                            Caption ="Summary Lists"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =6633
                                    Top =2777
                                    Width =1359
                                    Height =585
                                    FontWeight =700
                                    TabIndex =14
                                    ForeColor =8404992
                                    Name ="Button65"
                                    Caption ="Generate Event Lists"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =4500
                                    Top =3960
                                    Width =1125
                                    Height =737
                                    TabIndex =15
                                    ForeColor =8404992
                                    Name ="ProgramOfEvents"
                                    Caption ="Program of Events\015\012(4 Columns)"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =6993
                                    Top =3960
                                    Width =1125
                                    Height =737
                                    TabIndex =16
                                    ForeColor =8404992
                                    Name ="ProgramSummaryBut"
                                    Caption ="Program of Events Summary"
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
                                    Left =3060
                                    Top =912
                                    Width =5209
                                    Height =3908
                                    Name ="Box136"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =3173
                                    Top =735
                                    Width =4230
                                    Height =255
                                    BackColor =-2147483633
                                    Name ="Label137"
                                    Caption ="Generate lists for events satisfying this criteria:"
                                    FontName ="Tahoma"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =255
                                    Left =6469
                                    Top =1757
                                    Width =1694
                                    Height =1762
                                    Name ="Box138"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =6462
                                    Top =1587
                                    Width =750
                                    Height =255
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Label139"
                                    Caption ="List type:"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =3220
                                    Top =2115
                                    Width =915
                                    Height =255
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Label140"
                                    Caption ="Final status:"
                                    FontName ="Tahoma"
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =5700
                                    Top =3960
                                    Width =1125
                                    Height =737
                                    TabIndex =17
                                    ForeColor =8404992
                                    Name ="ProgOfEvents3Cols"
                                    Caption ="Program of Events\015\012(3 Columns)"
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
                            Left =2910
                            Top =510
                            Width =5355
                            Height =4305
                            Name ="Page127"
                            Caption ="Misc. Lists"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =255
                                    Left =3232
                                    Top =713
                                    Width =4266
                                    Height =847
                                    Name ="Box102"
                                End
                                Begin CheckBox
                                    Visible = NotDefault
                                    OverlapFlags =247
                                    Left =5621
                                    Top =596
                                    Name ="EntrySheet"
                                    ControlSource ="Rentry"
                                    DefaultValue ="False"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =3316
                                            Top =599
                                            Width =2070
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text97"
                                            Caption ="Generic Event Entry Sheets"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    Left =4129
                                    Top =937
                                    TabIndex =1
                                    BackColor =16777215
                                    BorderColor =12632256
                                    Name ="Field108"
                                    ControlSource ="Rhead1"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =3288
                                            Top =941
                                            Width =780
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text109"
                                            Caption ="Heading1:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    Left =4129
                                    Top =1220
                                    TabIndex =2
                                    BackColor =16777215
                                    BorderColor =12632256
                                    Name ="Field110"
                                    ControlSource ="Rhead2"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =3288
                                            Top =1224
                                            Width =780
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            Name ="Text111"
                                            Caption ="Heading2:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =6406
                                    Top =2494
                                    Width =1125
                                    Height =1005
                                    TabIndex =3
                                    ForeColor =8404992
                                    Name ="GenerateSpecial"
                                    Caption ="Generate Age / Gender / Team Lists"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =6249
                                    Top =4081
                                    Width =1125
                                    Height =570
                                    TabIndex =4
                                    ForeColor =8404992
                                    Name ="GenerateNameTags"
                                    Caption ="Generate Name Tags"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =6249
                                    Top =850
                                    Width =1095
                                    Height =600
                                    TabIndex =5
                                    ForeColor =8404992
                                    Name ="GenericListBut"
                                    Caption ="Generic Entry Sheets"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin TextBox
                                    OldBorderStyle =1
                                    OverlapFlags =255
                                    Left =5382
                                    Top =4365
                                    Width =636
                                    TabIndex =6
                                    BackColor =16777215
                                    BorderColor =12632256
                                    Name ="Text130"
                                    ControlSource ="NameTagFontSize"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =4265
                                            Top =4370
                                            Width =1065
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Label131"
                                            Caption ="Font Size:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =3244
                                    Top =3911
                                    Width =4366
                                    Height =851
                                    Name ="Box133"
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =3180
                                    Top =1920
                                    Width =2490
                                    Height =1785
                                    TabIndex =7
                                    Name ="Houses"
                                    SourceObject ="Form.Statistical Reports - Team SF"

                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            TextAlign =0
                                            Left =3180
                                            Top =1680
                                            Width =720
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Houses Label"
                                            Caption ="Houses"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="Houses_Label"
                                        End
                                    End
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

    Y = [Forms]![EnterCompetitors]![Sel_Event]
    x = RTrim([Forms]![EnterCompetitors]![Sex_DD])

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
     
  PleaseWaitMsg = "Generating event lists ..."
  DoCmd.RunMacro "ShowPleaseWait"
  DoCmd.RunCommand acCmdSaveRecord
  
  Dim Criteria As String, db As Database, Rs As Recordset
  Dim NewTitle As String
  Dim Filter As String

  Call GenerateFinalStatusFilter(Filter)
    
  Set db = DBEngine.Workspaces(0).Databases(0)

  Q = "SELECT DISTINCTROW EventType.ET_Code, EventType.Flag, EventType.R_Code "
  Q = Q & "FROM EventType WHERE (EventType.Flag = True AND EventType.Include = true) ORDER BY EventType.R_Code"

  Set Rs = db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.
    
  If Rs.BOF Then
    Response = MsgBox("No Events have been selected thus no report will be generated.", vbInformation)
  ElseIf Me![SummaryReport] = False And Me!Detailed = False Then
    Response = MsgBox("Specify the type of list you wish to generate.", vbInformation)
  Else
    Rs.MoveFirst
    
    Old_R_Code = -1

    Do Until Rs.EOF  ' Loop until no matching records.
        
        R_Code = Rs!R_Code
        If R_Code <> Old_R_Code Then
            If Me![SummaryReport] Then
                ReportName = DLookup("[SummaryReport]", "ReportTypes", "[R_Code] = " & R_Code)
                If Not IsNull(ReportName) Then
                    DoCmd.OpenReport ReportName, A_PREVIEW, , "[R_Code] = " & R_Code & " AND (" & Filter & ")"
                    DoCmd.Maximize
                End If
            End If
            
            If Me![Detailed] Then
                ReportName = DLookup("[Report]", "ReportTypes", "[R_Code] = " & R_Code)
                If Not IsNull(ReportName) Then
                    DoCmd.OpenReport ReportName, A_PREVIEW, , "[R_Code] = " & R_Code & " AND (" & Filter & ")"
                    DoCmd.Maximize
                End If
            End If

        End If

        Old_R_Code = R_Code

        Rs.MoveNext
    Loop
  End If
  Rs.Close
  
  
Exit_Button65_Click:
  DoCmd.RunMacro "ClosePleaseWait"
  DoCmd.OpenForm "ReportsPopUp"
  Exit Sub

Err_Button65_Click:
  If Err.Number = 2501 Then ' No Data error
    Resume Next
  Else
    MsgBox Error$
    Resume Exit_Button65_Click
  End If
    
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

Private Sub Flev_DD_DblClick(Cancel As Integer)

    Me![Flev_DD] = "*"

End Sub

Private Sub Form_Activate()
    
    DoCmd.SelectObject A_FORM, "Reports_Event", False
    DoCmd.Restore

End Sub

Private Sub Form_Load()

    Future_AfterUpdate
    Active_AfterUpdate
    Completed_AfterUpdate
    Promoted_AfterUpdate

End Sub

Private Sub Future_AfterUpdate()

    If Me![Future] = True Then
        Me![FutureEB] = 0
    Else
        Me![FutureEB] = 9
    End If

End Sub

Private Sub GenerateNameTags_Click()
On Error GoTo Err_GenerateNameTags_Click

  Dim DocName As String
  
  DoCmd.RunCommand acCmdSaveRecord
  
  DocName = "Name Tags"
  DoCmd.OpenReport DocName, A_PREVIEW
  DoCmd.Maximize

  DoCmd.OpenForm "ReportsPopUp"
  

Exit_GenerateNameTags_Click:
    Exit Sub

Err_GenerateNameTags_Click:
    MsgBox Error$
    Resume Exit_GenerateNameTags_Click
    
End Sub

Private Sub GenerateSpecial_Click()
On Error GoTo Err_GenerateSpecial_Click

  DoCmd.RunCommand acCmdSaveRecord
  
  Dim DocName As String

  DocName = "CompetitorList-ByTeamAge"
  DoCmd.OpenReport DocName, A_PREVIEW
  DoCmd.Maximize

  DoCmd.OpenForm "ReportsPopUp"
  

Exit_GenerateSpecial_Click:
    Exit Sub

Err_GenerateSpecial_Click:
    MsgBox Error$
    Resume Exit_GenerateSpecial_Click
    
End Sub

Private Sub Heat_EB_DblClick(Cancel As Integer)

    Me![Heat_EB] = "*"

End Sub

Private Sub ProgOfEvents3Cols_Click()
On Error GoTo ProgOfEvents3Cols_Click_Err

  DoCmd.RunCommand acCmdSaveRecord

  Dim stDocName As String, Filter As String
  
  Call GenerateFinalStatusFilter(Filter)
  
  stDocName = "Program of Events-3 Col"
  DoCmd.OpenReport stDocName, acPreview, , Filter
  DoCmd.Maximize

  DoCmd.OpenForm "ReportsPopUp"

ProgOfEvents3Cols_Click_Exit:
  On Error Resume Next
  Exit Sub

ProgOfEvents3Cols_Click_Err:
  Call DisplayErrMsg("ProgOfEvents3Cols_Click")
  Resume ProgOfEvents3Cols_Click_Exit

End Sub

Private Sub ProgramSummaryBut_Click()
On Error GoTo Err_ProgramSummaryBut_Click
  
  DoCmd.RunCommand acCmdSaveRecord
    
  Dim stDocName As String, Filter As String

  Call GenerateFinalStatusFilter(Filter)
    
  stDocName = "Program of Events-Summary"
  DoCmd.OpenReport stDocName, acPreview, , Filter
  DoCmd.Maximize

  DoCmd.OpenForm "ReportsPopUp"
  

Exit_ProgramSummaryBut_Click:
    Exit Sub

Err_ProgramSummaryBut_Click:
  If Err.Number <> 2501 Then ' not No data error
    MsgBox Err.Description
  End If
  Resume Exit_ProgramSummaryBut_Click
    
End Sub

Private Sub Promoted_AfterUpdate()

    If Me![Promoted] = True Then
        Me![PromotedEB] = 3
    Else
        Me![PromotedEB] = 9
    End If


End Sub
Private Sub RSfld_AfterUpdate()
    
    Forms![reports_event]![Event_DD] = "*"

End Sub

Private Sub Sel_Event_Change()

    Y = 1
    z = 2
    K = 3

    x = [Forms]![EnterCompetitors]![Sel_Event]

End Sub

Private Sub Sel_Event_DblClick(Cancel As Integer)


    Y = 1
    z = 2
    K = 3

    x = [Forms]![EnterCompetitors]![Sel_Event]
    Y = x + 1


End Sub

Private Sub Sex_DD_DblClick(Cancel As Integer)

    Me![Sex_DD] = "*"

End Sub

Private Sub ProgramOfEvents_Click()
On Error GoTo Err_ProgramOfEvents_Click

  DoCmd.RunCommand acCmdSaveRecord

  Dim stDocName As String, Filter As String
  
  Call GenerateFinalStatusFilter(Filter)
  
  stDocName = "Program of Events"
  DoCmd.OpenReport stDocName, acPreview, , Filter
  DoCmd.Maximize

  DoCmd.OpenForm "ReportsPopUp"
  

Exit_ProgramOfEvents_Click:
    Exit Sub

Err_ProgramOfEvents_Click:
  If Err.Number <> 2501 Then ' Not no data error
    MsgBox Err.Description
  End If
  Resume Exit_ProgramOfEvents_Click
    
End Sub
Private Sub GenericListBut_Click()
On Error GoTo Err_GenericListBut_Click

  Dim stDocName As String
  
  DoCmd.RunCommand acCmdSaveRecord
  
  stDocName = "EventResultsEntrySheet"
  DoCmd.OpenReport stDocName, acPreview
  DoCmd.Maximize

  DoCmd.OpenForm "ReportsPopUp"
  

Exit_GenericListBut_Click:
    Exit Sub

Err_GenericListBut_Click:
    MsgBox Err.Description
    Resume Exit_GenericListBut_Click
    
End Sub
Private Sub SaveRecord()
On Error GoTo Err_Command134_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Command134_Click:
    Exit Sub

Err_Command134_Click:
    MsgBox Err.Description
    Resume Exit_Command134_Click
    
End Sub

Private Sub GenerateFinalStatusFilter(ByRef Filter As String)
  
  Filter = "[Status]=" & Me!FutureEB & " OR "
  Filter = Filter & "[Status]=" & Me!ActiveEB & " OR "
  Filter = Filter & "[Status]=" & Me!CompletedEB & " OR "
  Filter = Filter & "[Status]=" & Me!PromotedEB

End Sub
