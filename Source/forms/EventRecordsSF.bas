Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridY =10
    Width =13429
    ItemSuffix =11
    Left =1530
    Top =4410
    Right =11520
    Bottom =7590
    HelpContextId =280
    RecSrcDt = Begin
        0x7e678ecfd0dae140
    End
    RecordSource ="SELECT DISTINCTROW Records.E_Code, Records.Surname, Records.Gname, Records.H_Cod"
        "e, Records.Comments, Records.nResult, Records.Result, EventType.Units, Records.D"
        "ate FROM (EventType INNER JOIN (Events INNER JOIN Records ON Events.E_Code = Rec"
        "ords.E_Code) ON EventType.ET_Code = Events.ET_Code) INNER JOIN Units ON EventTyp"
        "e.Units = Units.DisplayUnit ORDER BY IIf([Order]=\"ASC\",[nResult],1/[nResult]),"
        " Records.Surname, Records.Gname;"
    AfterUpdate ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =56
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =291
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =4755
                    Top =30
                    Width =681
                    TabIndex =7
                    Name ="Units"
                    ControlSource ="Units"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =56
                    Top =30
                    Width =1146
                    Name ="Surname"
                    ControlSource ="Surname"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =1246
                    Top =30
                    Width =1236
                    TabIndex =1
                    Name ="Gname"
                    ControlSource ="Gname"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    Left =3854
                    Top =30
                    Width =891
                    TabIndex =3
                    Name ="Record"
                    ControlSource ="Result"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    Left =7704
                    Width =726
                    TabIndex =5
                    Name ="nResult"
                    ControlSource ="nResult"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    Left =6911
                    Width =786
                    TabIndex =6
                    Name ="E_Code"
                    ControlSource ="E_Code"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =2080
                    Left =2520
                    Top =30
                    Width =1285
                    TabIndex =2
                    ColumnInfo ="\"\";\">\";\"\";\"\";\"10\";\"100\""
                    Name ="Field4"
                    ControlSource ="H_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="Select [H_Code],[H_NAme] From [House];"
                    ColumnWidths ="0;1830"
                    FontName ="Tahoma"
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5475
                    Top =30
                    Width =876
                    TabIndex =4
                    Name ="Date"
                    ControlSource ="Date"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6405
                    Top =15
                    Width =336
                    Height =263
                    TabIndex =8
                    Name ="DeleteBut"
                    Caption ="Command10"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadadadadadaadad00adad00adaddadad00ad00adada ,
                        0xadadad0000adadaddadadad00adadadaadadad0000adadaddadad00ad00adada ,
                        0xadad00adad00adaddadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub Field3_AfterUpdate()

End Sub

Private Sub Form_AfterUpdate()

    Forms![EventRecord].Refresh

End Sub

Private Sub Record_AfterUpdate()
    
  Dim res As String
  Dim Runit As String
  Dim nValu As String
  Dim success As Boolean
  
  res = Me![Record]
  
  If Not (IsNull(res)) Then
    
    nValu = ""
    Runit = Me![Units]
    Call Calculate_Results(res, nValu, Runit, success)

    Me![Record] = nValu
    Me![nResult] = res

  Else
    ' When the Result (time or distance or points) is set to NULL then
    ' set Numeric Result and Place to 0

    Me.[nRecord] = 0
    
  End If

End Sub

Private Sub Result_AfterUpdate()


End Sub

Private Sub DeleteBut_Click()
On Error GoTo Err_DeleteBut_Click

  Response = MsgBox("Are you sure you want to delete this record?", vbYesNo + vbCritical)
  If Response = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
  End If

Exit_DeleteBut_Click:
    DoCmd.SetWarnings True
    Exit Sub

Err_DeleteBut_Click:
    MsgBox Err.Description
    Resume Exit_DeleteBut_Click
    
End Sub
