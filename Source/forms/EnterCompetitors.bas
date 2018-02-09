Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =9694
    ItemSuffix =167
    Left =8295
    Top =2640
    Right =17985
    Bottom =9705
    HelpContextId =110
    RecSrcDt = Begin
        0xf2f778be6e4ae240
    End
    RecordSource ="EnterCompetitors"
    Caption ="Enter Results"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Line
            BorderLineStyle =0
            Width =1701
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
        Begin OptionGroup
            BorderLineStyle =0
            Width =1701
            Height =1701
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
        Begin FormHeader
            Height =1474
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    BackStyle =0
                    Left =831
                    Top =105
                    Width =3465
                    Height =256
                    ColumnOrder =1
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="ET_Code"
                    ControlSource ="ET_Des"
                    StatusBarText ="Every Event must fall under some event type"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =129
                            Top =105
                            Width =672
                            Height =256
                            FontSize =10
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text18"
                            Caption ="Event:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    Left =2140
                    Top =456
                    Width =300
                    Height =256
                    ColumnOrder =2
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="SexFld"
                    ControlSource ="Sex"
                    StatusBarText ="Male (M) / Female (F)"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =1695
                            Top =456
                            Width =422
                            Height =256
                            FontWeight =400
                            Name ="Text21"
                            Caption ="Sex:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    Left =831
                    Top =456
                    Width =825
                    Height =256
                    ColumnOrder =3
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="AgeFld"
                    ControlSource ="Age"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =1
                            Left =114
                            Top =456
                            Width =700
                            Height =256
                            FontSize =9
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text24"
                            Caption ="Age:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5151
                    Top =757
                    Width =555
                    Height =226
                    ColumnOrder =4
                    TabIndex =4
                    BackColor =16777215
                    BorderColor =12632256
                    HelpContextId =10150
                    Name ="Heat"
                    ControlSource ="Heat"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =4575
                            Top =735
                            Width =505
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text57"
                            Caption ="Heat:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    BackStyle =0
                    Left =3046
                    Top =456
                    Width =1245
                    Height =256
                    ColumnOrder =5
                    TabIndex =5
                    Name ="RecordUnits"
                    ControlSource ="=nz([Record],\"-\") & \" \" & [units]"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =2478
                            Top =456
                            Width =465
                            Height =256
                            FontSize =9
                            FontWeight =400
                            Name ="Text66"
                            Caption ="Rec.:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =5148
                    Top =225
                    Width =525
                    Height =256
                    ColumnOrder =6
                    TabIndex =6
                    BackColor =8421631
                    BorderColor =12632256
                    HelpContextId =10160
                    Name ="F_Lev"
                    ControlSource ="F_Lev"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =4560
                            Top =227
                            Width =525
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text81"
                            Caption ="Final:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    Left =5865
                    Top =750
                    Width =617
                    Height =256
                    ColumnOrder =7
                    TabIndex =7
                    BackColor =-2147483633
                    Name ="NoHeats"
                    FontName ="Tahoma"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    Left =5730
                    Top =742
                    Width =165
                    Height =256
                    FontSize =9
                    Name ="Text100"
                    Caption ="/"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    Left =5724
                    Top =225
                    Width =332
                    Height =256
                    ColumnOrder =8
                    TabIndex =8
                    BackColor =-2147483633
                    Name ="NoFinals"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =5100
                    Width =471
                    Height =170
                    ColumnOrder =9
                    TabIndex =9
                    BackColor =16744703
                    Name ="Field112"
                    ControlSource ="=Count([HE_Code])"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =1105
                    Left =6771
                    Top =225
                    Width =1105
                    Height =256
                    ColumnOrder =10
                    TabIndex =10
                    BackColor =8421504
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"20\""
                    Name ="FinalStatus"
                    ControlSource ="Status"
                    RowSourceType ="Table/Query"
                    RowSource ="Select [StatusID],[Status] From [FinalStatus];"
                    ColumnWidths ="0;855"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =7727
                    Top =792
                    ColumnOrder =11
                    TabIndex =11
                    BorderColor =12632256
                    Name ="HeatComplete"
                    ControlSource ="Completed"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =6876
                            Top =735
                            Width =720
                            Height =256
                            FontSize =7
                            FontWeight =400
                            Name ="Text117"
                            Caption ="Complete:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =4506
                    Top =105
                    Width =3498
                    Height =981
                    BackColor =12440319
                    Name ="Box118"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =247
                    Left =6068
                    Top =225
                    Width =585
                    Height =256
                    FontWeight =400
                    Name ="Text120"
                    Caption ="Status:"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =8841
                    Top =1190
                    ColumnOrder =0
                    BorderColor =12632256
                    Name ="AllNames"
                    ControlSource ="AllNames"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =2
                            Left =8160
                            Top =690
                            Width =1455
                            Height =435
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text130"
                            Caption ="Select from all Names"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =87
                    BackStyle =0
                    Left =840
                    Top =807
                    Width =1170
                    Height =256
                    ColumnOrder =13
                    FontWeight =700
                    TabIndex =13
                    Name ="PtScale"
                    ControlSource ="PtScale"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =120
                            Top =811
                            Width =720
                            Height =256
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text133"
                            Caption ="Pt Scale:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    Left =3083
                    Top =801
                    Width =1200
                    Height =256
                    ColumnOrder =12
                    FontWeight =700
                    TabIndex =12
                    Name ="Lane_Cnt"
                    ControlSource ="Lane_Cnt"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =2065
                            Top =803
                            Width =915
                            Height =256
                            FontWeight =400
                            Name ="Text128"
                            Caption ="Total Lanes:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =3456
                    Width =1005
                    Height =170
                    ColumnOrder =14
                    TabIndex =14
                    BackColor =16744703
                    Name ="Record"
                    ControlSource ="Record"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =4530
                    Width =495
                    Height =170
                    ColumnOrder =15
                    TabIndex =15
                    BackColor =16744703
                    Name ="nRecord"
                    ControlSource ="nRecord"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"

                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =3
                    Left =4629
                    Top =621
                    Width =3315
                    Name ="Line157"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =850
                    Width =801
                    Height =165
                    ColumnOrder =17
                    TabIndex =17
                    BackColor =16744703
                    Name ="E_Code"
                    ControlSource ="E_Code"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =215
                            Width =765
                            Height =165
                            BackColor =16744703
                            Name ="Text94"
                            Caption ="E_Code"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    Left =2607
                    Width =801
                    Height =165
                    ColumnOrder =18
                    TabIndex =18
                    BackColor =16744703
                    Name ="Field95"
                    ControlSource ="HE_Code"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =1635
                            Top =7
                            Width =885
                            Height =165
                            BackColor =16744703
                            Name ="Text96"
                            Caption ="HE_Code"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =8355
                    Top =313
                    Width =1065
                    ColumnOrder =16
                    TabIndex =16
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="EventNumber"
                    ControlSource ="E_Number"
                    StatusBarText ="Age Group for event - ie. 13 year boys=13; 21 year girls =21; Open Boys=0"
                    FontName ="Tahoma"
                    ControlTipText ="The event number for this event."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =2
                            Left =8355
                            Top =30
                            Width =1065
                            Height =240
                            FontWeight =400
                            BackColor =16744703
                            Name ="Text71"
                            Caption ="Event #:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =2616
                    Top =1202
                    ColumnOrder =19
                    TabIndex =19
                    BorderColor =12632256
                    Name ="PlacesAcrossAllHeats"
                    ControlSource ="PlacesAcrossAllHeats"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =105
                            Top =1185
                            Width =2430
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Label161"
                            Caption ="Places allocated across all heats:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =7731
                    Top =1202
                    ColumnOrder =20
                    TabIndex =20
                    BorderColor =12632256
                    Name ="DontOverridePlaces"
                    ControlSource ="DontOverridePlaces"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =5955
                            Top =1185
                            Width =1695
                            Height =285
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="DontOverridePlacesLabel"
                            Caption ="Don't override places:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =5331
                    Top =1202
                    ColumnOrder =21
                    TabIndex =21
                    BorderColor =12632256
                    Name ="EffectsRecords"
                    ControlSource ="EffectsRecords"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =3135
                            Top =1185
                            Width =2115
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Label165"
                            Caption ="Heat Effects Event Records:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =87
                    Top =1455
                    Width =9666
                    BorderColor =12632256
                    Name ="Line166"
                    HorizontalAnchor =2
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5587
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =135
                    Top =90
                    Width =7928
                    Height =5437
                    TabIndex =6
                    Name ="EC_Subform"
                    SourceObject ="Form.Enter Competitors Subform1"
                    LinkChildFields ="E_Code;heat;F_Lev"
                    LinkMasterFields ="E_Code;heat;F_Lev"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8271
                    Top =3141
                    Width =1247
                    Height =504
                    FontSize =7
                    FontWeight =400
                    TabIndex =3
                    Name ="MaintainCompetitorsBut"
                    Caption ="Maintain Competitors"
                    OnClick ="[Event Procedure]"
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
                    Left =8271
                    Top =2565
                    Width =1247
                    Height =504
                    FontSize =7
                    FontWeight =400
                    TabIndex =2
                    Name ="CalculatePlacesBut"
                    Caption ="Calculate Places"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8235
                    Top =4971
                    Width =1247
                    Height =504
                    FontSize =8
                    TabIndex =5
                    Name ="DoneBut"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
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
                    Left =8271
                    Top =1989
                    Width =1247
                    Height =504
                    FontSize =7
                    FontWeight =400
                    TabIndex =1
                    Name ="EnterResultsInPlaceOrderBut"
                    Caption ="Enter Results in Place Order"
                    OnClick ="[Event Procedure]"
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
                    Left =8235
                    Top =4395
                    Width =1247
                    Height =504
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    HelpContextId =110
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
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =8198
                    Top =199
                    Width =1401
                    Height =1648
                    Name ="OrderBy"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =0
                            Left =8303
                            Top =79
                            Width =799
                            Height =255
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text139"
                            Caption ="Order by:"
                            FontName ="Tahoma"
                            HorizontalAnchor =1
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8473
                            Top =447
                            OptionValue =1
                            Name ="LaneOrder"
                            HorizontalAnchor =1

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    Left =8746
                                    Top =419
                                    Width =570
                                    Height =240
                                    FontWeight =400
                                    Name ="field1111"
                                    Caption ="Lane"
                                    FontName ="Tahoma"
                                    HorizontalAnchor =1
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8473
                            Top =765
                            OptionValue =2
                            Name ="Button143"
                            HorizontalAnchor =1

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    Left =8746
                                    Top =737
                                    Width =630
                                    Height =240
                                    FontWeight =400
                                    Name ="Text144"
                                    Caption ="Name"
                                    FontName ="Tahoma"
                                    HorizontalAnchor =1
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8473
                            Top =1083
                            OptionValue =3
                            Name ="Button145"
                            HorizontalAnchor =1

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    Left =8746
                                    Top =1055
                                    Width =630
                                    Height =240
                                    FontWeight =400
                                    Name ="Text146"
                                    Caption ="Place"
                                    FontName ="Tahoma"
                                    HorizontalAnchor =1
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8476
                            Top =1458
                            OptionValue =4
                            Name ="Button155"
                            HorizontalAnchor =1

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    TextAlign =1
                                    Left =8749
                                    Top =1430
                                    Width =630
                                    Height =240
                                    FontWeight =400
                                    Name ="Text156"
                                    Caption ="Not"
                                    FontName ="Tahoma"
                                    HorizontalAnchor =1
                                End
                            End
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8307
                    Top =3877
                    Height =202
                    TabIndex =7
                    Name ="QuickTab"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="True"
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =8540
                            Top =3849
                            Width =900
                            Height =255
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Label159"
                            Caption ="Quick Tab"
                            FontName ="Tahoma"
                            HorizontalAnchor =1
                        End
                    End
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

Option Explicit

Dim UpdateFinalStatus As Variant
Dim CalculatePlaces As Variant
Dim DontEditPromotionFinals As Variant
                                      
Dim AllPlacesFilled As Variant
Dim AllResultsFilled As Variant

' Form Dimensions
Dim lMinHeight As Long
Dim lMinWidth As Long

Private Sub AllNames_AfterUpdate()
On Error GoTo AllNames_AfterUpdate_Err

    PleaseWaitMsg = "Retrieving appropriate competitors ..."
    DoCmd.RunMacro "ShowPleaseWait"

    If Not ([AllNames]) Then
      Me![EC_Subform].Form![Fname].RowSource = GenerateAgeFilter(Me![AgeFld], Me.Sex)
      Me![EC_Subform].Form![Fname].Requery
       
    Else
      Me![EC_Subform].Form![Fname].RowSource = GenerateSexFilter(Me.Sex)
      Me![EC_Subform].Form![Fname].Requery
       
    End If
    
AllNames_AfterUpdate_Exit:
    DoCmd.RunMacro "ClosePleaseWait"
    Exit Sub
    
AllNames_AfterUpdate_Err:
  MsgBox "An error has occurred in 'AllNames_AfterUpdate': " & Err.Description, vbCritical
  GoTo AllNames_AfterUpdate_Exit
  
End Sub

Private Sub DontOverridePlaces_AfterUpdate()

    If Me.DontOverridePlaces Then
      Me.DontOverridePlacesLabel.ForeColor = vbRed
    Else
      Me.DontOverridePlacesLabel.ForeColor = vbBlack
    End If
    

End Sub

Private Sub EnterResultsInPlaceOrderBut_Click()

On Error GoTo EnterResultsInPlaceOrderBut_Click_Err

    Dim Criteria As String, db As Database, Rs As Recordset
    Dim MyDb As Database, MySet As Recordset, Q As Variant, X As Variant, i As Variant, Place As Variant, Lane As Variant
    Dim NewTitle As String, Criteria1 As Variant, ECrs As Recordset
    Dim Success As Boolean
    
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set Rs = db.OpenRecordset("Temporary Results-Place Order", dbOpenDynaset)   ' Create dynaset.
    Set ECrs = db.OpenRecordset("CompEvents", dbOpenDynaset)   ' Create dynaset.
    
    'Stop
    
    X = Me![EC_Subform].Form![Count]
    
    If X >= 1 Then
    
        Q = "DELETE DISTINCTROW [Temporary Results-Place Order].Place FROM [Temporary Results-Place Order]"
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True
    
        Criteria = "[E_Code]=" & Me![E_Code] & " AND [Heat]=" & Me![Heat] & " AND [F_Lev]=" & Me![F_Lev]
         
        'If x = 0 Then
        '    x = DCount("[PIN]", "CompEvents", "[E_Code]=" & Me![E_Code] & " AND [F_Lev]=" & Me![F_Lev] & " AND [Heat]=" & Me!Heat)
        'End If

        ECrs.FindFirst Criteria
        i = 1
        While Not ECrs.NoMatch
            Rs.AddNew
            Rs!Place = i
            Rs!AvailableLanes = ECrs!Lane
            Rs.Update
            ECrs.FindNext Criteria
            i = i + 1
        Wend
        
        'q = "UPDATE DISTINCTROW [Temporary Results-Place Order] SET [Temporary Results-Place Order].Lane = Null, [Temporary Results-Place Order].Results = Null"
    
        'DoCmd SetWarnings False
        'DoCmd RunSQL q
        'DoCmd SetWarnings True
        
        GlobalCancel = True
        DoCmd.OpenForm "Temporary Results-Place Order", , , , , acDialog
      
      'Stop

      If Not GlobalCancel Then
        
        PleaseWaitMsg = "Updating competitor placings and results ..."
        DoCmd.RunMacro "ShowPleaseWait"
        
        Set Rs = db.OpenRecordset("Temporary Results-Place Order", dbOpenDynaset)   ' Create dynaset.
    
        Dim nValu As String, Result As String, Runit As String
    
        Rs.MoveFirst
        Do Until Rs.EOF  ' Loop until no matching records.
            Place = Rs!Place
            Lane = Rs!Lane
            
            If IsNull(Rs!Results) Then
                Result = "0"
            Else
                Result = Rs!Results
            End If
            Runit = Me![Units]
            Call Calculate_Results(Result, nValu, Runit, Success)

            Q = "UPDATE DISTINCTROW CompEvents SET CompEvents.Place = " & Place & ", CompEvents.Result = """ & nValu & """, CompEvents.nResult = " & Result
            Q = Q & " WHERE CompEvents.E_Code=" & Me![E_Code] & " And CompEvents.F_Lev = " & Me![F_Lev] & " And CompEvents.Heat = " & Me![Heat] & " AND CompEvents.Lane= " & Lane
            
            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
            
            Rs.MoveNext
    
        Loop
    
        Rs.Close
            
        Criteria1 = "[HE_Code] = " & Me![HE_Code] & " AND [Place] = 0"
        If IsNull(DLookup("[HE_Code]", "EnterCompetitorsSF", Criteria1)) Then
            ' All places have been allocated so heat is complete
            Forms![EnterCompetitors]![HeatComplete].Value = True
        End If
    
        ' Determine Points
        CalculatePoints ("POINTS")
    
      End If
    Else
        MsgBox ("There are no competitors enrolled in this event.")
    End If

EnterResultsInPlaceOrderBut_Click_Exit:
  DoCmd.RunMacro "ClosePleaseWait"
  Exit Sub
  
EnterResultsInPlaceOrderBut_Click_Err:
  MsgBox ("An error has occured in [EnterResultsInPlaceOrderBut_Click]: " & Err.Description)
  GoTo EnterResultsInPlaceOrderBut_Click_Exit
End Sub

Private Sub Button47_Click()
    
    Dim X As Variant

    X = DCount("[ET_Code]", "Ent_Comp_Filter")
    If X < 1 Then
        MsgBox ("No events match the given criteria.")
    Else
        DoCmd.RunMacro "ApplyFilter"
    End If

End Sub

Private Sub DoneBut_Click()
    On Error GoTo Err_DoneBut_Click

    DoCmd.Close

Exit_DoneBut_Click:
    Exit Sub

Err_DoneBut_Click:
    MsgBox Error$
    Resume Exit_DoneBut_Click
    
End Sub

Private Sub Form_Close()
    On Error GoTo Err_Form_Close

    ' Checks if all heats have been been marked Complete, if so set Event level to Completed
    Call SetCurrentFinal(Me![E_Code])
    
    If UpdateFinalStatus Then
        Call SetAllFinalToSameValue([E_Code], [F_Lev], [FinalStatus])
    End If
    

    ' Update event list when closing.
    ' Need to catch error in case form not open.
    Forms!CompEventsSummary!Summary.Requery
    
Exit_Form_Close:
    Exit Sub
    
Err_Form_Close:
    'MsgBox Error$
    Resume Exit_Form_Close
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    lMinHeight = frmHeight(Me)
    lMinWidth = Me.Width
End Sub

Private Sub Form_Resize()
    If Not m_blResize Then Call glrMinWindowSize(Me, lMinHeight, lMinWidth, False)
End Sub

Private Sub MaintainCompetitorsBut_Click()
On Error GoTo Err_MaintainCompetitorsBut

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "CompetitorsSummary"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog

Exit_MaintainCompetitorsBut:
    Exit Sub

Err_MaintainCompetitorsBut:
    MsgBox Error$
    Resume Exit_MaintainCompetitorsBut
    
End Sub

Private Sub CalculatePlacesBut_Click()

    Dim Criteria As String, MyDb As Database, MySet As Recordset
    Dim NewTitle As String, Response As Variant, OK As Variant
        
    'If Not [HeatComplete] Then
    '    Response = MsgBox("It appears that the heat has not yet been completed (some competitors do not have results).  Do you still want to to calculate places?", vbYesNo + vbInformation + vbDefaultButton1, "Heat not complete.")
    'Else
    '    Response = vbYes
    'End If
    
    'If Response = vbYes Then
    '    If AllPlacesFilled Then
    '        If AllResultsFilled Then
    '            ' ERROR MSG FIX
    '            Response = MsgBox("Places will be calculated from the results that have been entered.  Calculating places for this event will overide existing places.  Do you want to continue?", vbYesNo + vbInformation + vbDefaultButton1, "Overwrite Exisiting Places")
    '        ElseIf Not IsNull(DLookup("[nResult]", "CompEvents", "CompEvents.E_Code=" & [E_Code] & " AND CompEvents.Heat= " & [Heat] & " AND CompEvents.F_Lev= " & [F_Lev] & " AND [nResult]>0")) Then
    '            ' Some Results
    '            'MsgBox ("All places have been entered but not all results.  Places will be left unchanged.")
    '            Response = vbNo
    '        Else
    '            ' There are no results so do nothing because all the places have already been entered.
    '            Response = vbNo
    '        End If
    '    Else ' Assume that AllResultsFilled must be true since we are executing this routine
    '        If IsNull(DLookup("[Place]", "CompEvents", "CompEvents.E_Code=" & [E_Code] & " AND CompEvents.Heat= " & [Heat] & " AND CompEvents.F_Lev= " & [F_Lev] & " AND [Place] <> 0")) Then
    '            ' No Places have been entered so no warning
    '            Response = vbYes
    '        Else
    '            ' Some places have been filled
    '            Response = MsgBox("Some places have already been entered.  New places will be calculated from the results that have been entered.  Calculating places for this event will overwrite existing places.  Do you want to continue?", vbYesNo + vbInformation + vbDefaultButton1, "Overwrite Exisiting Places")
    '
    '        End If
    '    End If

     '   If Response = vbYes Then
     
            PleaseWaitMsg = "Updating places ..."
            DoCmd.RunMacro "ShowPleaseWait"
            Call CalculatePoints("PLACE")
            DoCmd.RunMacro "ClosePleaseWait"

      '  End If
    'End If

End Sub

Private Sub CalculatePoints(Ty)
On Error GoTo CalculatePoints_Err

    ' If TY = Place then the Place will be determined and saved
    ' If TY = Points then the Place will NOT be determined and will be left as is

    Dim Criteria As String
    Dim CRrs As Recordset  ' (Competitor Results Recordset)
    
    Dim Points As Variant, qSQL As Variant, UN As Variant, PointScale As Variant
    Dim PL As Variant, BMark As Variant, Res1 As Variant, Res2 As Variant, i As Variant, j As Variant
    Dim PSrs As Recordset
    
    qSQL = "SELECT DISTINCTROW CompEvents.E_Code, CompEvents.Heat, CompEvents.F_Lev, CompEvents.nResult, CompEvents.Result, CompEvents.Place, CompEvents.Points "
    qSQL = qSQL & "FROM CompEvents "
    qSQL = qSQL & "WHERE ((CompEvents.E_Code=" & [E_Code] & ") AND "
    qSQL = qSQL & "(CompEvents.F_Lev =" & [F_Lev] & ")"
    
    If Not Me.PlacesAcrossAllHeats Then 'Excludes the Heat filter if places are to be calculated across all heats
      qSQL = qSQL & " AND (CompEvents.Heat= " & [Heat] & ")"
    End If
    
    qSQL = qSQL & ") "
    qSQL = qSQL & "AND ((CompEvents.nResult<>0) OR (CompEvents.Result=""PARTICIPATE"")) "
    qSQL = qSQL & "ORDER BY CompEvents.nResult "
    
    UN = UCase([Units])
    If UN = "M" Or UN = "KM" Or UN = "PTS" Then
        qSQL = qSQL & "DESC"
    End If

    Set CRrs = CurrentDb.OpenRecordset(qSQL, dbOpenDynaset)   ' Create dynaset.
    
    Criteria = "E_Code = " & [E_Code]
    Criteria = Criteria & " AND Heat = " & Me.Heat & " AND F_Lev = " & Me![F_Lev]
    
    PointScale = DLookup("[PtScale]", "Heats", Criteria)
    Set PSrs = CurrentDb.OpenRecordset("SELECT * FROM [PointsScale] WHERE [PtScale]=""" & PointScale & """")

    ''CRrs.FindFirst Criteria    ' Find first occurrence.
    
    PL = 1
    
    Do Until CRrs.EOF '' CRrs.NoMatch  ' Loop until no matching records.
        
      'If two or more competitors get the same result then places must be 'moved down' accordingly
      '  Find how many competitors got the same place and set their places.
      
      If CRrs!Result = "PARTICIPATE" Then
        CRrs.Edit
        CRrs!Points = Nz(DMin("[Points]", "PointsScale", "[PtScale]=""" & PointScale & """"), 0)
        CRrs!Place = Null
        CRrs.Update
        CRrs.MoveNext
        
      ElseIf CRrs!Result = "FOUL" Then
        CRrs.Edit
        CRrs!Points = Nz(DMin("[Points]", "PointsScale", "[PtScale]=""" & PointScale & """"), 0)
        CRrs!Place = Null
        CRrs.Update
        CRrs.MoveNext
      Else
      
        BMark = CRrs.Bookmark         ' Obtain bookmark.
        Res1 = CRrs!nResult
        CRrs.MoveNext
        
        ''CRrs.FindNext Criteria
        
        If Not CRrs.EOF Then Res2 = CRrs!nResult
  
        i = 1
  
        While (Res1 = Res2) And (Not CRrs.EOF) ''(Not CRrs.NoMatch)
            i = i + 1
            ''CRrs.FindNext Criteria
            CRrs.MoveNext
            If Not CRrs.EOF Then Res2 = CRrs!nResult
  
        Wend
  
        CRrs.Bookmark = BMark
  
        For j = 1 To i
          CRrs.Edit          ' Enable editing.
          If Ty = "PLACE" Then
              CRrs!Place = PL
          End If
          
          If IsNull(CRrs!Place) Then
              Points = Null
          Else
          
            'PSrs.FindFirst "[PtScale]=""" & PointScale & """" & " AND [Place]=" & PL
            PSrs.FindFirst "[Place]=" & PL
            
            If Not PSrs.NoMatch Then
              Points = Nz(PSrs!Points, 0)
            Else
              Points = 0
            End If
            'Points = DLookup("[Points]", "PointsScale", "[PtScale]=""" & PointScale & """" & " AND [Place]=" & PL)
            'If IsNull(Points) Then
            '    Points = 0
            'End If
            
          End If
  
          CRrs!Points = Points
  
          CRrs.Update        ' Save changes.
          CRrs.MoveNext
          
        Next j
        
        PL = PL + i
        
      End If
      
    Loop

    CRrs.Close
    
    Me.EC_Subform.Requery
    
CalculatePoints_Exit:
  Exit Sub
  
CalculatePoints_Err:
  MsgBox "An error has occurred in [CalculatePoints]: " & Err.Description, vbCritical
  Resume CalculatePoints_Exit
  
End Sub


Private Sub xCalculatePointsGivenPlace()

    Dim Q As Variant
    ' NOT USED

    Q = "UPDATE DISTINCTROW Heats INNER JOIN CompEvents ON (Heats.Heat = CompEvents.Heat) AND (Heats.F_Lev = CompEvents.F_Lev) AND (Heats.E_Code = CompEvents.E_Code) SET CompEvents.Points = DeterminePoints([Place],[PtScale]) "
    Q = Q & "WHERE (CompEvents.E_Code=" & Me![E_Code] & ") AND (CompEvents.F_Lev=" & Me![F_Lev] & ") AND (CompEvents.Heat=" & Me![Heat] & ")"
    
    'DoCmd SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True

End Sub

Private Sub EC_Subform_Enter()

' On Error Resume Next
        
  DontEditPromotionFinalsMessage = True
  Me![CalculatePlacesBut].enabled = False
  GlobalChange = False
  GlobalPlaceChange = False
  CalculatePlaces = False
  Me![EC_Subform].Form.Refresh

End Sub

Private Sub EC_Subform_Exit(Cancel As Integer)
On Error Resume Next
  Dim Q As String, Response As Integer

  If Not GlobalChange Then Exit Sub
  
  If GlobalPlaceChange And Not Me.DontOverridePlaces Then
    Q = "You have made manual changes to competitor places which could be over written when the Sports Administrator "
    Q = Q & "attempts to calculate places automatically.  " & LFCR & LFCR
    Q = Q & "Do you wish to stop this from happening by ticking the "
    Q = Q & "'Dont override place' check box? " & LFCR
    Q = Q & "(Ticking this box means places will no longer be calculated for this heat automatically)"
    Response = MsgBox(Q, vbYesNo + vbDefaultButton1 + vbQuestion)
    If Response = vbYes Then Me.DontOverridePlaces = True
  End If
    
  Dim Criteria1 As Variant, Criteria2 As Variant

  'Stop
  PleaseWaitMsg = "Checking competitor results and places ..."
  DoCmd.RunMacro "ShowPleaseWait"

  AllPlacesFilled = False
  AllResultsFilled = False

  'Ensure that heat exists
  If Not IsNull(DLookup("[HE_Code]", "EnterCompetitorsSF", "[HE_Code] = " & Me![HE_Code])) Then
  
    Criteria1 = "[HE_Code] = " & Me![HE_Code] & " AND nz([Place]) = 0"
    Criteria2 = "[HE_Code] = " & Me![HE_Code] & " AND nz([nResult]) = 0"

    If IsNull(DLookup("[HE_Code]", "EnterCompetitorsSF", Criteria1)) Then
      ' All places have been allocated so heat is complete

      Me.HeatComplete = True
      AllPlacesFilled = True
    End If
    
    If IsNull(DLookup("[HE_Code]", "EnterCompetitorsSF", Criteria2)) Then
      ' All results have been allocated so heat is complete

      Me.HeatComplete = True
      AllResultsFilled = True
    End If

    GlobalVariable = True
    If Me.EffectsRecords Then Call CheckIfRecordBroken(Me![E_Code], Me![Heat], Me![F_Lev])
              
    If Not Me.DontOverridePlaces Then 'OK to override places if they exist
      Call CalculatePlacesBut_Click
      CalculatePlaces = False
    End If
        
    If Me![Lane_Cnt] > 0 Then ' Must be a limited lanes event
        If Not IsNull(DLookup("[Lane]", "CompEvents", "[E_Code]=" & Me![E_Code] & " AND [F_Lev]=" & Me![F_Lev] & " AND [Heat]=" & Me![Heat] & " AND [Lane]=0")) Then
            'Stop
            If DLookup("[ShowNoAllocatedLane]", "[Misc-EnterCompetitorEvents]") = True Then
                MsgBox "One or more competitors have not been assigned lanes (lane=0) which means that the competitor(s) will not be displayed in any reports.  This is due to the house or school that the competitor belongs to not having had sufficient lanes allocated to it.  You can check the house/school lane allocation from the Maintain Houses form.", vbInformation
            End If
        End If
    End If
  
  End If

  Me![CalculatePlacesBut].enabled = True

  Me.EC_Subform.Form.RecordsetClone.MoveFirst
  Me.EC_Subform.Form.Bookmark = Me.EC_Subform.Form.RecordsetClone.Bookmark
  
  DoCmd.RunMacro "ClosePleaseWait"
  GlobalChange = False
  
End Sub

Private Sub Enum_DblClick(Cancel As Integer)

    Me![Enum] = "*"

End Sub

Private Sub Event_DD_DblClick(Cancel As Integer)

    Me![Event_DD] = "*"

End Sub


Private Sub FinalCompleted()

    If IsNull(DLookup("[Completed]", "Heats", "[Completed] = NO And [E_Code] = " & [Forms]![EnterCompetitors]![E_Code] & " And [F_Lev] = " & [Forms]![EnterCompetitors]![Fle] & " and [Heat] <> " & [Forms]![EnterCompetitors]![Heat])) Then
        Forms![EnterCompetitors]![Fcomp] = True
    Else
        Forms![EnterCompetitors]![Fcomp] = False
    End If
         
End Sub

Private Sub FinalStatus_AfterUpdate()
    
    UpdateFinalStatus = True
    If FinalStatus > 1 Then
        Me![HeatComplete] = True
    Else
        Me![HeatComplete] = False
    End If
    Call Form_Current

End Sub

Private Sub FinalStatus_BeforeUpdate(Cancel As Integer)

    Dim Response As Variant

    Response = MsgBox("Changing the status of a final is not recommended.  Are you sure you want to change the status?", vbExclamation + vbYesNo + vbDefaultButton2)
    If Response = vbNo Then
        Cancel = True
    End If
        
End Sub

Private Sub FirstRec_Click()
On Error GoTo Err_FirstRec_Click


    DoCmd.GoToRecord , , A_FIRST

Exit_FirstRec_Click:
    Exit Sub

Err_FirstRec_Click:
    MsgBox Error$
    Resume Exit_FirstRec_Click
    
End Sub

Private Sub Flevel_DblClick(Cancel As Integer)

    Me![Flevel] = "*"

End Sub

Private Sub Form_AfterUpdate()

 '  Moved to Close action to prevent warning about data edited while form open.
 
   ' Call SetCurrentFinal(Me![E_Code])
    
  '  If UpdateFinalStatus Then
  '      Call SetAllFinalToSameValue([E_Code], [F_Lev], [FinalStatus])
  '  End If


End Sub

Private Sub Form_Current()
 On Error Resume Next

    CalculatePlaces = True ' Set this flag to calculate places next time the CalculatePlaces button is pushed
    Me.EC_Subform.Form![Fname].Requery
    Me.NoHeats = NumberOfHeats()
    Me.NoFinals = NumberOfFinals()

    If Me![Lane_Cnt] = 0 Then
        Me![Button131].visible = False
        Me![EC_Subform].Form![Lane].visible = False
        Me![EC_Subform].Form![LaneTXT].visible = False
        Me![LaneOrder].visible = False
        Me![OrderBy] = 2

    Else
        Me![Button131].visible = True
        Me![EC_Subform].Form![Lane].visible = True
        Me![EC_Subform].Form![LaneTXT].visible = True
        Me![LaneOrder].visible = True
    End If

    
    If [Status] = evStatus.Future Then ' Future
        Me![F_Lev].BackColor = White
        Me![FinalStatus].BackColor = White

    ElseIf [Status] = evStatus.Current Then ' Current
        Me![F_Lev].BackColor = LightBlue
        Me![FinalStatus].BackColor = LightBlue

    ElseIf [Status] = evStatus.Completed Then ' Completed
        Me![F_Lev].BackColor = LightRed
        Me![FinalStatus].BackColor = LightRed

    ElseIf [Status] = evStatus.Promoted Then ' Promoted
        Me![F_Lev].BackColor = DarkGrey
        Me![FinalStatus].BackColor = DarkGrey

    End If

    Dim db As Database
    Set db = CurrentDb
    
    'Me![EC_Subform].Form![Fname].RowSource = GenerateAgeFilter([Forms]![EnterCompetitors]![AgeFld])
    
    'db.QueryDefs("EnterCompetitorsSF-CompetitorList").SQL = GenerateAgeFilter([Forms]![EnterCompetitors]![AgeFld])
    'Me![EC_Subform].Form![Fname].RowSource = "EnterCompetitorsSF-CompetitorList"
    'Me![EC_Subform].Form![Fname].RowSource = GenerateAgeFilter(Me.AgeFld, Me.Sex)
    'Me![EC_Subform].Form![Fname].Requery
    
    Call AllNames_AfterUpdate
    
    Call DontOverridePlaces_AfterUpdate
    
    'Forms![main form name]![subform control name].Form![control name]

Form_Current_Exit:
    Exit Sub

Form_Current_Err:
    MsgBox (Error$)
    GoTo Form_Current_Exit

End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Err

    DoCmd.RunMacro "ClosePleaseWait"
    UpdateFinalStatus = False
    If SportsViewModule Then
      Me.EventNumber.Locked = True
      Me.EventNumber.enabled = False
      
      Me.Heat.Locked = True
      Me.Heat.enabled = False
      
      Me.FinalStatus.Locked = True
      Me.FinalStatus.enabled = False
      
      Me.F_Lev.Locked = True
      Me.F_Lev.enabled = False
      
      Me.HeatComplete.Locked = True
      Me.HeatComplete.enabled = False
      
      Me.Heat.Locked = True
      Me.Heat.enabled = False
      
      Me.DontOverridePlaces.Locked = True
      Me.DontOverridePlaces.enabled = False
      
      Me.PlacesAcrossAllHeats.Locked = True
      Me.PlacesAcrossAllHeats.enabled = False
      
      Me.AllNames.Locked = True
      Me.AllNames.enabled = False
      
      Me.EC_Subform.Form.AllowEdits = False
      Me.EC_Subform.Form.AllowAdditions = False
      Me.EC_Subform.Form.AllowDeletions = False
      
      Me.EC_Subform.Locked = True
      
      Me.EC_Subform.Form.DeleteCompetitorBut.visible = False
      Me.EC_Subform.Form.Memo.visible = False
      
      Me.EnterResultsInPlaceOrderBut.visible = False
      Me.CalculatePlacesBut.visible = False
      Me.MaintainCompetitorsBut.visible = False
      Me.Help.visible = False
      
    End If
    
    ' Set Focus to Name Entry
    Me.EC_Subform.SetFocus

Form_Load_Exit:
  Exit Sub
  
Form_Load_Err:
  MsgBox "An error has occurred in [Form_LoadForm_Load]: " & Err.Description, vbCritical
  
End Sub

Private Function NumberOfFinals()
    
    Dim Criteria As Variant

    Criteria = "[E_Code] = " & [E_Code] & " AND [Heat] = 1"
    NumberOfFinals = DCount("[F_Lev]", "NumberOfHeats", Criteria) - 1

End Function

Private Function NumberOfHeats()

    Dim Criteria As Variant

    Criteria = "[E_Code] = " & [E_Code] & " AND [F_Lev] = " & [F_Lev]
    NumberOfHeats = DCount("[F_Lev]", "NumberOfHeats", Criteria)
    
End Function

Private Sub OrderBy_AfterUpdate()

    PleaseWaitMsg = "Modifying display order ..."
    DoCmd.RunMacro "ShowPleaseWait"
    
    If Me![OrderBy] = 1 Then ' Order by Lane
        Me![EC_Subform].Form.RecordSource = "EnterCompetitorsSF"

    ElseIf Me![OrderBy] = 2 Then ' Order by name
        Me![EC_Subform].Form.RecordSource = "EnterCompetitorsSf-Ordered by Name"

    ElseIf Me![OrderBy] = 3 Then ' Order by Lane
        Me![EC_Subform].Form.RecordSource = "EnterCompetitorsSf-Ordered by Place"
    
    ElseIf Me![OrderBy] = 4 Then ' Order by Lane
        Me![EC_Subform].Form.RecordSource = "EnterCompetitorsSf-Not Ordered"

    End If

    Me![EC_Subform].Requery

    DoCmd.RunMacro "ClosePleaseWait"
        
End Sub

Private Sub QuickTab_AfterUpdate()

  Me.EC_Subform.Form.DeleteCompetitorBut.TabStop = Not Me.QuickTab
  Me.EC_Subform.Form.Lane.TabStop = Not Me.QuickTab
  Me.EC_Subform.Form.Memo.TabStop = Not Me.QuickTab
  Me.EC_Subform.Form.Points.TabStop = Not Me.QuickTab
  
  
End Sub

Private Sub RecordUnits_DblClick(Cancel As Integer)

    DoCmd.OpenForm "EventRecordHistory", , , "[E_Code]=" & Me![E_Code], , acDialog

End Sub

Private Sub SetAllFinalToSameValue(E_Code, F_Lev, FinalStatus)

    Dim Q As Variant

    Q = "UPDATE DISTINCTROW Heats SET Heats.Status = " & FinalStatus
    Q = Q & " WHERE Heats.E_Code=" & E_Code & " AND Heats.F_Lev=" & F_Lev

    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
    
End Sub
