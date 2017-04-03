Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    WhatsThisButton = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =20
    Width =8674
    ItemSuffix =115
    Left =600
    Top =90
    Right =11145
    Bottom =8280
    HelpContextId =60
    Filter ="[ET_Code] = 24"
    RecSrcDt = Begin
        0x6cca3c042dc7e140
    End
    RecordSource ="SELECT DISTINCTROW EventType.ET_Code, EventType.ET_Des, EventType.Units, EventTy"
        "pe.Lane_Cnt, EventType.R_Code, ReportTypes.Desc, EventType.EntrantNum, EventType"
        ".Include FROM ReportTypes INNER JOIN EventType ON ReportTypes.R_Code = EventType"
        ".R_Code ORDER BY EventType.ET_Des;"
    Caption ="Add Event Wizard"
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
            Height =6689
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7380
                    Top =6165
                    Width =1134
                    Height =397
                    FontSize =8
                    FontWeight =400
                    Name ="Close"
                    Caption ="Finish"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =195
                    Top =120
                    Width =8340
                    Height =5850
                    TabIndex =1
                    Name ="TabCtl"
                    FontName ="Tahoma"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =330
                            Top =525
                            Width =8070
                            Height =5316
                            Name ="Page79"
                            Caption ="General Details"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =560
                                    Top =4921
                                    Width =3405
                                    Height =675
                                    FontWeight =400
                                    Name ="Label93"
                                    Caption ="How are results recorded for this event.  (100m sprint would be recorded in seco"
                                        "nds, Long Jump in meters etc.)"
                                    FontName ="Tahoma"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =223
                                    Left =390
                                    Top =4303
                                    Width =3710
                                    Height =1531
                                    Name ="Box94"
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =567
                                    Top =949
                                    Width =3165
                                    Height =245
                                    HelpContextId =400
                                    Name ="ET_Des"
                                    ControlSource ="ET_Des"
                                    StatusBarText ="Event Description - ie. 200m; 100m Hurdles; High Jump- No sex / age specifics"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =567
                                            Top =683
                                            Width =1140
                                            Height =245
                                            Name ="Text19"
                                            Caption ="Description:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =567
                                    Top =3082
                                    Width =930
                                    Height =245
                                    TabIndex =1
                                    HelpContextId =420
                                    Name ="Lane_Cnt"
                                    ControlSource ="Lane_Cnt"
                                    StatusBarText ="Lane / Competitor Count"
                                    AfterUpdate ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =567
                                            Top =2814
                                            Width =2550
                                            Height =290
                                            Name ="Text23"
                                            Caption ="Lane Count (0 if unlimited)"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    ColumnCount =2
                                    ListWidth =3402
                                    Left =567
                                    Top =2061
                                    Width =3175
                                    Height =245
                                    TabIndex =2
                                    HelpContextId =410
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="R_Code"
                                    ControlSource ="R_Code"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT ReportTypes.R_Code, ReportTypes.Desc, ReportTypes.EventReport FROM Report"
                                        "Types WHERE ((ReportTypes.EventReport=Yes)) ORDER BY ReportTypes.Desc;"
                                    ColumnWidths ="0;3452"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =567
                                            Top =1806
                                            Width =1260
                                            Height =245
                                            Name ="Text39"
                                            Caption ="Report Style:"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    ColumnCount =2
                                    ListWidth =1210
                                    Left =560
                                    Top =4630
                                    Width =1600
                                    Height =245
                                    TabIndex =3
                                    HelpContextId =430
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="Units"
                                    ControlSource ="Units"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT DISTINCTROW Units.DisplayUnit, Units.Unit FROM Units;"
                                    ColumnWidths ="0;903"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =215
                                            TextAlign =1
                                            Left =560
                                            Top =4376
                                            Width =930
                                            Height =245
                                            Name ="Text46"
                                            Caption ="Units"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =0
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =4471
                                    Top =4653
                                    Width =930
                                    Height =240
                                    TabIndex =4
                                    HelpContextId =440
                                    Name ="EntrantNum"
                                    ControlSource ="EntrantNum"
                                    StatusBarText ="Lane / Competitor Count"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =4478
                                            Top =4385
                                            Width =3285
                                            Height =245
                                            Name ="Text66"
                                            Caption ="Number of Entrants from each Team"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =3916
                                    Top =969
                                    TabIndex =5
                                    Name ="Include"
                                    ControlSource ="Include"
                                    DefaultValue ="Yes"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            TextAlign =1
                                            Left =4172
                                            Top =913
                                            Width =2220
                                            Height =240
                                            BackColor =-2147483633
                                            Name ="Text72"
                                            Caption ="Include Event in Carnival"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =567
                                    Top =1253
                                    Width =7425
                                    Height =285
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Label87"
                                    Caption ="The event description should not be gender or age specific. Correct examples: 10"
                                        "0m Sprint, High Jump"
                                    FontName ="Tahoma"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =397
                                    Top =629
                                    Width =7625
                                    Height =1003
                                    Name ="Box88"
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =567
                                    Top =2346
                                    Width =7365
                                    Height =210
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Label89"
                                    Caption ="The marshalling list style varies depending on the event type.  Select the corre"
                                        "ct style for this event."
                                    FontName ="Tahoma"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =397
                                    Top =1725
                                    Width =7640
                                    Height =928
                                    Name ="Box90"
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =567
                                    Top =3369
                                    Width =7290
                                    Height =840
                                    FontWeight =400
                                    BackColor =-2147483633
                                    Name ="Label91"
                                    Caption ="Specify the number of lanes available for this event.  This value is used when p"
                                        "romoting competitors from one final level into the next.  This value determines "
                                        "how many competitors are allowed in each race.  Enter 0 for events where the lan"
                                        "e count is not applicable (eg. 800m, 1500m, most field events etc)"
                                    FontName ="Tahoma"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =247
                                    Left =397
                                    Top =2751
                                    Width =7625
                                    Height =1453
                                    Name ="Box92"
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =255
                                    Left =4313
                                    Top =4310
                                    Width =3710
                                    Height =1531
                                    Name ="Box95"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =4478
                                    Top =4925
                                    Width =3405
                                    Height =795
                                    FontWeight =400
                                    Name ="Label96"
                                    Caption ="This value is used when generating carnival disks.  Specify the number of compet"
                                        "itors from each team that are allowed to compete in each heat"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =330
                            Top =525
                            Width =8070
                            Height =5310
                            Name ="Page80"
                            Caption ="Event Divisions"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =411
                                    Top =858
                                    Width =3535
                                    Height =4932
                                    Name ="ET_Sub1"
                                    SourceObject ="Form.EventTypeSub1"
                                    LinkChildFields ="ET_Code"
                                    LinkMasterFields ="ET_Code"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =255
                                            TextAlign =0
                                            Left =390
                                            Top =600
                                            Width =1950
                                            Height =240
                                            Name ="Text34"
                                            Caption ="Age \\ Gender Divisions"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =4095
                                    Top =630
                                    Width =4185
                                    Height =840
                                    FontWeight =400
                                    Name ="Label97"
                                    Caption ="Enter the various age \\ gender divisions you require for this event.  For examp"
                                        "le, if you are conducting a secondary carnival then you age divisions would be s"
                                        "imilar to:\015\012"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =4195
                                    Top =1470
                                    Width =3415
                                    Height =265
                                    Name ="Label100"
                                    Caption ="13_U, 14, 15, 16, 17_O, OPEN"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =4095
                                    Top =1785
                                    Width =4185
                                    Height =1845
                                    FontWeight =400
                                    Name ="Label101"
                                    Caption ="If you have both male and female competitors you would create an age division fo"
                                        "r each gender.\015\012\015\012You can use any alpha-numeric text to represent yo"
                                        "ur age divisions.  However, by specifying the divisions as shown above (with the"
                                        " age first) the Sports Administrator will show only the appropriate competitors."
                                        "\015\012\015\012Note that you can use:\015\012\015\012"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =4200
                                    Top =3645
                                    Width =4020
                                    Height =600
                                    Name ="Label102"
                                    Caption ="_U      : to represent 'and under'\015\012_O      : to represent 'and over'\015\012"
                                        "OPEN  : for ALL competitors"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =4050
                                    Top =4365
                                    Width =4185
                                    Height =1320
                                    FontWeight =400
                                    Name ="Label104"
                                    Caption ="RECORDS:\015\012\015\012After you have created your age-divisions you can enter "
                                        "the records for an event by double-clicking in the record field of the division "
                                        "or by clicking once in the division and pushing the 'Edit Record for Selected Di"
                                        "vision' button.\015\012"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =330
                            Top =525
                            Width =8070
                            Height =5310
                            Name ="Page81"
                            Caption ="Event Heats"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =570
                                    Top =855
                                    Width =7620
                                    Height =1785
                                    Name ="Final Level SF"
                                    SourceObject ="Form.Final_Lev_Sub"
                                    LinkChildFields ="ET_Code"
                                    LinkMasterFields ="ET_Code"
                                    EventProcPrefix ="Final_Level_SF"

                                    Begin
                                        Begin Label
                                            BackStyle =0
                                            OverlapFlags =255
                                            TextAlign =0
                                            Left =574
                                            Top =615
                                            Width =1980
                                            Height =240
                                            Name ="Final Level SF Label"
                                            Caption ="Heats and Final Setup"
                                            FontName ="Tahoma"
                                            EventProcPrefix ="Final_Level_SF_Label"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    TextFontFamily =34
                                    Left =4500
                                    Top =5430
                                    Width =3684
                                    Height =397
                                    FontSize =8
                                    FontWeight =400
                                    TabIndex =1
                                    HelpContextId =530
                                    Name ="Command108"
                                    Caption ="Specific Help on Setting up Heats and Finals"
                                    OnClick ="Open Help"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =570
                                    Top =2775
                                    Width =7590
                                    Height =2580
                                    FontWeight =400
                                    Name ="Label114"
                                    Caption ="The Sports Administrator can handle any number of final levels.  Number them con"
                                        "secutively starting at 0 and continuing in whole numbers (ie. 0, 1, 2, 3 and so "
                                        "on).\015\012\015\012Start numbering your final levels from 0  (grand-final).  En"
                                        "ter the number of heats that will be offered in that final level (for field even"
                                        "ts it is usually 1, track events can be any number).  Specify the Pointscale you"
                                        " desire and set the PromotionMethod to 'None'.  \015\012\015\012Extra final leve"
                                        "ls are needed only if you will be promoting competitors from, say, a semi-final "
                                        "into a grand-final.  If so, enter 1 for your final level, enter the number of he"
                                        "ats in that final level, the Pointscale used for the final level, the PromotionM"
                                        "ethod, and whether you wish to promote competitors by result or place.\015\012\015"
                                        "\012Add extra final levels if you need them.  "
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =330
                            Top =525
                            Width =8070
                            Height =5310
                            Name ="Page84"
                            Caption ="Lane Promotion"
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =396
                                    Top =886
                                    Width =1935
                                    Height =4350
                                    Name ="Lane Promotion Allocation SF"
                                    SourceObject ="Form.Lane Promotion Allocation SF"
                                    LinkChildFields ="ET_Code"
                                    LinkMasterFields ="ET_Code"
                                    EventProcPrefix ="Lane_Promotion_Allocation_SF"

                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =255
                                    TextAlign =0
                                    Left =660
                                    Top =630
                                    Width =675
                                    Height =240
                                    Name ="Label106"
                                    Caption ="Place"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =255
                                    TextAlign =0
                                    Left =1382
                                    Top =623
                                    Width =750
                                    Height =240
                                    Name ="Label107"
                                    Caption ="Lane"
                                    FontName ="Tahoma"
                                End
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =247
                                    TextAlign =1
                                    Left =2607
                                    Top =878
                                    Width =4455
                                    Height =3405
                                    FontWeight =400
                                    Name ="Label109"
                                    Caption ="If you have a number of final-levels in this event and so will be promoting from"
                                        " one final to the next you should setup what lanes competitors will be promoted "
                                        "into.  Normally the fastest competitors are given the 'best' lanes in the next r"
                                        "ace the compete in (usually the middle lanes).\015\012\015\012Enter the place th"
                                        "e competitor attained on the left and the lane which he or she receives in the n"
                                        "ext race on the right.\015\012\015\012There should be the same number of lane pr"
                                        "omotion entries as there are available lanes."
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4308
                    Top =6166
                    Width =1134
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="NextBut"
                    Caption ="Next >>"
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
                    Left =3089
                    Top =6166
                    Width =1134
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    Name ="PrevBut"
                    Caption ="<< Previous"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =822
                    Top =6264
                    TabIndex =4
                    Name ="ET_Code"
                    ControlSource ="ET_Code"
                    FontName ="Tahoma"

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

Private Sub Close_Click()
On Error GoTo Err_Close_But_Click

  If AutomaticallyCreateHeatsAndFinals(Me!ET_Code, True) Then
    GoTo Close_But_Click_Now
  Else
    Response = MsgBox("Do you wish to fix the problem?", vbYesNo + vbInformation + vbDefaultButton1, "Resolve the problem?")
    If Response = vbNo Then GoTo Close_But_Click_Now
  End If

Exit_Close_But_Click:
    Exit Sub

Err_Close_But_Click:
    MsgBox Error$
    GoTo Exit_Close_But_Click
    
Close_But_Click_Now:
  Response = MsgBox("Your new event has been created.  Push OK to view the details of your new event.  Then push 'Done' to go back to the 'Event Summary' form where you can add more events.", vbInformation)
  DoCmd.Close
  Exit Sub
End Sub

Private Sub Form_Load()

  Call TabCtl_Change
  
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

Private Sub NextBut_Click()

  Me!TabCtl = Me!TabCtl + 1
  
End Sub

Private Sub PrevBut_Click()
  
  Me!TabCtl = Me!TabCtl - 1

End Sub

Private Sub TabCtl_Change()

  If Me!TabCtl = 0 Then
    Me!NextBut.visible = True
    Me!NextBut.SetFocus
    Me!PrevBut.visible = False
  ElseIf Me!TabCtl = 3 Then
    Me!PrevBut.visible = True
    Me!PrevBut.SetFocus
    Me!NextBut.visible = False
  Else
    Me!PrevBut.visible = True
    Me!NextBut.visible = True
  End If
  
End Sub
