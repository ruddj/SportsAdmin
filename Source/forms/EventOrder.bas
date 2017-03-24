Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
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
    GridX =20
    GridY =20
    Width =10318
    ItemSuffix =109
    Left =-18570
    Top =2730
    Right =-8250
    Bottom =11430
    HelpContextId =250
    RecSrcDt = Begin
        0x24586aaaf1e5e140
    End
    RecordSource ="Misc-OrderEvents"
    Caption ="Maintain Event Order"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    OnResize ="[Event Procedure]"
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
            CanGrow = NotDefault
            Height =0
            BackColor =128
            Name ="FormHeader1"
        End
        Begin Section
            CanGrow = NotDefault
            Height =6996
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ListWidth =1134
                    Left =8576
                    Top =427
                    Width =1195
                    Height =255
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="O1"
                    ControlSource ="First"
                    RowSourceType ="Value List"
                    RowSource ="\"Event #\";\"Description\";\"Age\";\"Sex\";\"Final Level\";\"Heat\""
                    ColumnWidths ="1134"
                    DefaultValue ="\"Event #\""
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =7710
                            Top =427
                            Width =735
                            Height =240
                            FontWeight =400
                            Name ="Text30"
                            Caption ="First:"
                            FontName ="Arial"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =8050
                    Top =84
                    Width =1695
                    Height =285
                    FontWeight =400
                    Name ="Text50"
                    Caption ="Order the events by:"
                    FontName ="Arial"
                    HorizontalAnchor =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ListWidth =1134
                    Left =8576
                    Top =767
                    Width =1195
                    Height =255
                    TabIndex =1
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="o2"
                    ControlSource ="Second"
                    RowSourceType ="Value List"
                    RowSource ="\"Event #\";\"Description\";\"Age\";\"Sex\";\"Final Level\";\"Heat\""
                    ColumnWidths ="1134"
                    DefaultValue ="\"Description\""
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =7710
                            Top =767
                            Width =735
                            Height =240
                            FontWeight =400
                            Name ="dfgdf"
                            Caption ="Second:"
                            FontName ="Arial"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ListWidth =1134
                    Left =8576
                    Top =1108
                    Width =1195
                    Height =255
                    TabIndex =2
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="o3"
                    ControlSource ="Third"
                    RowSourceType ="Value List"
                    RowSource ="\"Event #\";\"Description\";\"Age\";\"Sex\";\"Final Level\";\"Heat\""
                    ColumnWidths ="1134"
                    DefaultValue ="\"Age\""
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =7710
                            Top =1108
                            Width =750
                            Height =240
                            FontWeight =400
                            Name ="dgdg"
                            Caption ="Third:"
                            FontName ="Arial"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ListWidth =1134
                    Left =8576
                    Top =1448
                    Width =1195
                    Height =255
                    TabIndex =3
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="o4"
                    ControlSource ="Fourth"
                    RowSourceType ="Value List"
                    RowSource ="\"Event #\";\"Description\";\"Age\";\"Sex\";\"Final Level\";\"Heat\""
                    ColumnWidths ="1134"
                    DefaultValue ="\"Sex\""
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =7710
                            Top =1448
                            Width =750
                            Height =240
                            FontWeight =400
                            Name ="Text75"
                            Caption ="Fourth:"
                            FontName ="Arial"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8505
                    Top =2970
                    Width =1149
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    Name ="Refresh"
                    Caption ="Refresh Display"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =288
                    Top =72
                    Width =558
                    Height =226
                    FontWeight =400
                    BackColor =16777215
                    Name ="Text12"
                    Caption ="#"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =960
                    Top =72
                    Width =1815
                    Height =210
                    FontWeight =400
                    BackColor =16777215
                    Name ="Text15"
                    Caption ="Description"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2887
                    Top =77
                    Width =345
                    Height =210
                    FontWeight =400
                    BackColor =16777215
                    Name ="Text16"
                    Caption ="Sex"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3457
                    Top =77
                    Width =360
                    Height =210
                    FontWeight =400
                    BackColor =16777215
                    Name ="Text19"
                    Caption ="Age"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4032
                    Top =79
                    Width =420
                    Height =225
                    FontWeight =400
                    BackColor =16777215
                    Name ="Text20"
                    Caption ="Final"
                    FontName ="Arial"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4589
                    Top =72
                    Width =420
                    Height =210
                    FontWeight =400
                    BackColor =16777215
                    Name ="Text80"
                    Caption ="Heat"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =9060
                    Top =6330
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =5
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ListWidth =1134
                    Left =8576
                    Top =1808
                    Width =1195
                    Height =255
                    TabIndex =6
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="o5"
                    ControlSource ="Fifth"
                    RowSourceType ="Value List"
                    RowSource ="\"Event #\";\"Description\";\"Age\";\"Sex\";\"Final Level\";\"Heat\""
                    ColumnWidths ="1134"
                    DefaultValue ="\"Final Level\""
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =7710
                            Top =1808
                            Width =750
                            Height =240
                            FontWeight =400
                            Name ="Text83"
                            Caption ="Fifth:"
                            FontName ="Arial"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ListWidth =1134
                    Left =8576
                    Top =2168
                    Width =1195
                    Height =255
                    TabIndex =7
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="o6"
                    ControlSource ="Sixth"
                    RowSourceType ="Value List"
                    RowSource ="\"Event #\";\"Description\";\"Age\";\"Sex\";\"Final Level\";\"Heat\""
                    ColumnWidths ="1134"
                    DefaultValue ="\"Heat\""
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =7710
                            Top =2168
                            Width =750
                            Height =240
                            FontWeight =400
                            Name ="Text85"
                            Caption ="Sixth:"
                            FontName ="Arial"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8490
                    Top =5100
                    Width =1149
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    ForeColor =8404992
                    Name ="Default Order"
                    Caption ="Default Order"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    EventProcPrefix ="Default_Order"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =8160
                    Top =3763
                    Width =260
                    Height =210
                    TabIndex =9
                    BorderColor =12632256
                    Name ="AutoNumber"
                    ControlSource ="AutoRenumber"
                    DefaultValue ="0"
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =8424
                            Top =3720
                            Width =1185
                            Height =240
                            FontWeight =400
                            BackColor =16777215
                            Name ="Text91"
                            Caption ="Auto Renumber"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8475
                    Top =5670
                    Width =1149
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =10
                    ForeColor =8404992
                    Name ="SlideUp"
                    Caption ="Slide Event Numbers Up"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7875
                    Top =6345
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =11
                    HelpContextId =250
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Arial"
                    HorizontalAnchor =1

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9921
                    Top =427
                    Width =170
                    Height =200
                    TabIndex =12
                    Name ="O1C"
                    ControlSource ="FirstOrder"
                    DefaultValue ="Yes"
                    HorizontalAnchor =1

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =9809
                    Top =147
                    Width =390
                    Height =225
                    FontSize =7
                    FontWeight =400
                    Name ="Text95"
                    Caption ="ASC"
                    FontName ="Arial"
                    HorizontalAnchor =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9921
                    Top =768
                    Width =170
                    Height =200
                    TabIndex =13
                    Name ="o2C"
                    ControlSource ="SecondOrder"
                    DefaultValue ="Yes"
                    HorizontalAnchor =1

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9921
                    Top =1108
                    Width =170
                    Height =200
                    TabIndex =14
                    Name ="O3C"
                    ControlSource ="ThirdOrder"
                    DefaultValue ="Yes"
                    HorizontalAnchor =1

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9921
                    Top =1448
                    Width =170
                    Height =200
                    TabIndex =15
                    Name ="O4C"
                    ControlSource ="FourthOrder"
                    DefaultValue ="Yes"
                    HorizontalAnchor =1

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9921
                    Top =1845
                    Width =170
                    Height =200
                    TabIndex =16
                    Name ="O5C"
                    ControlSource ="FifthOrder"
                    DefaultValue ="Yes"
                    HorizontalAnchor =1

                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =9921
                    Top =2185
                    Width =170
                    Height =200
                    TabIndex =17
                    Name ="O6C"
                    ControlSource ="SixthOrder"
                    DefaultValue ="Yes"
                    HorizontalAnchor =1

                End
                Begin Line
                    OverlapFlags =93
                    SpecialEffect =3
                    Left =7766
                    Top =3628
                    Width =2400
                    Name ="Line102"
                    HorizontalAnchor =1
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5115
                    Top =90
                    Width =2400
                    Height =225
                    FontWeight =400
                    BackColor =16777215
                    Name ="Label103"
                    Caption ="Event Time"
                    FontName ="Arial"
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =8362
                    Top =2598
                    Width =260
                    Height =210
                    TabIndex =18
                    BorderColor =12632256
                    Name ="ShowAll"
                    ControlSource ="SwitchSex"
                    DefaultValue ="0"
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =8629
                            Top =2551
                            Width =1230
                            Height =240
                            FontWeight =400
                            BackColor =16777215
                            Name ="Label105"
                            Caption ="Show All Events"
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =165
                    Top =368
                    Width =7406
                    Height =6437
                    TabIndex =19
                    Name ="EventOrderSub"
                    SourceObject ="Form.EventOrderSub"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =3
                    Left =7756
                    Width =0
                    Height =6996
                    Name ="Line101"
                    HorizontalAnchor =1
                    LayoutCachedLeft =7756
                    LayoutCachedWidth =7756
                    LayoutCachedHeight =6996
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =8160
                    Top =4063
                    Width =260
                    Height =210
                    TabIndex =20
                    BorderColor =12632256
                    Name ="SingleClickOption"
                    DefaultValue ="False"
                    HorizontalAnchor =1

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =8430
                            Top =4020
                            Width =1710
                            Height =240
                            FontWeight =400
                            BackColor =16777215
                            Name ="Label107"
                            Caption ="Use single click option "
                            HorizontalAnchor =1
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8490
                    Top =4515
                    Width =1149
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =21
                    ForeColor =128
                    Name ="ClearAllNumbers"
                    Caption ="Clear all numbers"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    HorizontalAnchor =1

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

' Form Dimensions
Dim lMinHeight As Long
Dim lMinWidth As Long

Private Sub ClearAllNumbers_Click()

  Response = MsgBox("Are you sure you want to clear all event numbers?", vbQuestion + vbYesNo + vbDefaultButton2)
  
  If Response = vbYes Then
    DoCmd.RunSQL "UPDATE Heats SET Heats.E_Number = Null"
  End If
  
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

Private Sub Default_Order_Click()

    Me![O1] = "Event #"
    Me![o2] = "Description"
    Me![o3] = "Age"
    Me![o4] = "Sex"
    Me![o5] = "Final Level"
    Me![o6] = "Heat"
    Me![SwitchSex] = -1
    
End Sub

Private Sub Form_Open(Cancel As Integer)

    lMinHeight = frmHeight(Me)
    lMinWidth = Me.Width

    ' Update the Total Points field in Competitor's Table
    '
    'For Each Competitor
    '   Total Points =
    '           Find each event competior is in
    '           Determine point scale and place for each event
    '           Find points allocated to place and add to total points

    Q = "UPDATE DISTINCTROW Heats SET Heats.E_Number = Null WHERE Heats.E_Number=0"
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True
    
    Call OrderEvents


End Sub

Private Sub Form_Resize()
    If Not m_blResize Then Call glrMinWindowSize(Me, lMinHeight, lMinWidth, True)
End Sub

Private Sub OrderEvents()

    Squery = "SELECT DISTINCTROW Heats.E_Number, EventType.ET_Des, Events.Age, Events.Sex, Heats.F_lev, Heats.Heat, Heats.E_Time "
    Squery = Squery & " FROM EventType INNER JOIN (Events INNER JOIN Heats ON Events.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.ET_Code"
    If Not Me!ShowAll Then Squery = Squery & " WHERE EventType.Include = True "
    Squery = Squery & " ORDER BY "

    For i = 1 To 6
        Select Case i
         Case 1
             oType = Me![O1]
             OrderAsc = Me![O1C]
         Case 2
             oType = Me![o2]
             OrderAsc = Me![o2C]
         Case 3
             oType = Me![o3]
             OrderAsc = Me![O3C]
         Case 4
             oType = Me![o4]
             OrderAsc = Me![O4C]
         Case 5
             oType = Me![o5]
             OrderAsc = Me![O5C]
         Case 6
             oType = Me![o6]
             OrderAsc = Me![O6C]

        End Select
        
       Select Case oType
        Case "Event #"
            Squery = Squery & "Heats.E_Number"
            If OrderAsc Then
                Squery = Squery & " ASC"
            Else
                Squery = Squery & " DESC"
            End If
        
        
        Case "Description"
            Squery = Squery & "EventType.ET_Des"
            If OrderAsc Then
                Squery = Squery & " ASC"
            Else
                Squery = Squery & " DESC"
            End If
        
        Case "Sex"
            Squery = Squery & "Events.Sex"
            If OrderAsc Then
                Squery = Squery & " ASC"
            Else
                Squery = Squery & " DESC"
            End If
        
'            If Me![SwitchSex] Then
'                Squery = Squery & " DESC"
'            End If

        Case "Age"
            Squery = Squery & "Events.Age"
            If OrderAsc Then
                Squery = Squery & " ASC"
            Else
                Squery = Squery & " DESC"
            End If
        
        
        Case "Final Level"
            Squery = Squery & "Heats.F_Lev"
            If OrderAsc Then
                Squery = Squery & " ASC"
            Else
                Squery = Squery & " DESC"
            End If
        
        
        Case "Heat"
            Squery = Squery & "Heats.Heat"
            If OrderAsc Then
                Squery = Squery & " ASC"
            Else
                Squery = Squery & " DESC"
            End If
        
       End Select

        If i < 6 Then
            Squery = Squery & ", "
        Else
            'Squery = Squery & ";"
        End If
    Next i

    Me.EventOrderSub.Form.RecordSource = Squery
    Me.EventOrderSub.Form.Requery
     
End Sub

Private Sub Refresh_Click()

    Call OrderEvents

End Sub


Private Sub SlideUp_Click()

    
    Dim Criteria As String, rs As Recordset, Enumm As Variant
    
    Set rs = CurrentDb.OpenRecordset("select * from heats order by [E_Number]", dbOpenDynaset)   ' Create dynaset.

    Enumm = 1
    rs.MoveFirst
    Do Until rs.EOF
        If Not IsNull(rs![E_Number]) Then
            rs.Edit
            rs![E_Number] = Enumm
            rs.Update
            Enumm = Enumm + 1
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Me![EventOrderSub].Requery
    
End Sub
