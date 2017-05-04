Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    Width =7056
    ItemSuffix =30
    Left =2265
    Top =3195
    Right =9075
    Bottom =7830
    RecSrcDt = Begin
        0x386f898110cde140
    End
    Caption ="Copy Competitors from One Event to Another"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Rectangle
            BorderLineStyle =0
        End
        Begin CommandButton
            TextFontFamily =2
            BorderLineStyle =0
        End
        Begin TextBox
            BorderLineStyle =0
        End
        Begin ComboBox
            BorderLineStyle =0
        End
        Begin Section
            Height =5400
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =147
                    Top =360
                    Width =3168
                    Height =3840
                    Name ="Box3"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =600
                    Width =720
                    Height =210
                    Name ="Text2"
                    Caption ="Event:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =1200
                    Width =885
                    Height =210
                    Name ="Text4"
                    Caption ="Age:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =1800
                    Width =885
                    Height =210
                    Name ="Text5"
                    Caption ="Gender:"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    OverlapFlags =215
                    ColumnCount =2
                    ListWidth =2380
                    Left =435
                    Top =840
                    Width =2380
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="FmET_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT EventType.ET_Code, EventType.ET_Des FROM EventType WHERE (((EventType.Inc"
                        "lude)=Yes)) ORDER BY EventType.ET_Des;"
                    ColumnWidths ="0;2130"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =2640
                    TabIndex =3
                    BorderColor =12632256
                    Name ="FmFinalLev"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =3240
                    TabIndex =4
                    BorderColor =12632256
                    Name ="FmHeat"
                    Format ="Fixed"
                    BeforeUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =2400
                    Width =885
                    Height =210
                    Name ="Text11"
                    Caption ="Final Level:"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =435
                    Top =3000
                    Width =885
                    Height =210
                    Name ="Text12"
                    Caption ="Heat:"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =435
                    Top =1440
                    TabIndex =1
                    BorderColor =12632256
                    Name ="FmAge"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =215
                    Left =435
                    Top =2040
                    TabIndex =2
                    BorderColor =12632256
                    Name ="FmSex"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =144
                    Top =120
                    Width =1155
                    Height =210
                    FontWeight =700
                    Name ="Text15"
                    Caption ="From Event"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =93
                    Left =3462
                    Top =360
                    Width =3168
                    Height =3840
                    Name ="Box16"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =3750
                    Top =600
                    Width =720
                    Height =210
                    Name ="Text17"
                    Caption ="Event:"
                    FontName ="Tahoma"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =3750
                    Top =1200
                    Width =885
                    Height =210
                    Name ="Text18"
                    Caption ="Age:"
                    FontName ="Tahoma"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =3750
                    Top =1800
                    Width =885
                    Height =210
                    Name ="Text19"
                    Caption ="Gender:"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    OverlapFlags =215
                    ColumnCount =2
                    ListWidth =2380
                    Left =3750
                    Top =840
                    Width =2380
                    TabIndex =5
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="ToET_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT EventType.ET_Code, EventType.ET_Des FROM EventType WHERE (((EventType.Inc"
                        "lude)=Yes)) ORDER BY EventType.ET_Des;"
                    ColumnWidths ="0;2130"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =3750
                    Top =1440
                    TabIndex =6
                    BorderColor =12632256
                    Name ="ToAge"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =3750
                    Top =2040
                    TabIndex =7
                    BorderColor =12632256
                    Name ="ToSex"
                    FontName ="Tahoma"

                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =3750
                    Top =2400
                    Width =885
                    Height =210
                    Name ="Text23"
                    Caption ="Final Level:"
                    FontName ="Tahoma"
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    Left =3750
                    Top =3000
                    Width =885
                    Height =210
                    Name ="Text24"
                    Caption ="Heat:"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =3750
                    Top =2640
                    TabIndex =8
                    BorderColor =12632256
                    Name ="ToFinalLev"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =3750
                    Top =3240
                    TabIndex =9
                    BorderColor =12632256
                    Name ="ToHeat"
                    FontName ="Tahoma"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =3459
                    Top =120
                    Width =1155
                    Height =210
                    FontWeight =700
                    Name ="Text27"
                    Caption ="To Event"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4648
                    Top =4299
                    Width =1980
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =10
                    ForeColor =32768
                    Name ="CopyCompetitors"
                    Caption ="Copy Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =288
                    Top =4800
                    Width =1290
                    Height =465
                    FontWeight =400
                    TabIndex =11
                    Name ="CloseBut"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub CloseBut_Click()
On Error GoTo Err_CloseBut_Click


    DoCmd.Close

Exit_CloseBut_Click:
    Exit Sub

Err_CloseBut_Click:
    MsgBox Error$
    Resume Exit_CloseBut_Click
    
End Sub

Private Sub CopyCompetitors_Click()
On Error GoTo Err_CopyCompetitors_Click

  Dim db As Database, Frs As Recordset, Trs As Recordset, Retval As Variant
  Dim FailedI As Integer, msg As Variant, success As Variant
  
  If MsgBox("Are you sure you want to copy competitors from one event to another?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
  
  PleaseWaitMsg = "Copying competitors into specified event.  Please wait ..."
  DoCmd.RunMacro "ShowPleaseWait"
  
    ' For each competitor in FmEvent
    ' if Heat, F_Lev, AGe, Sex event exists in ToEventType
    '   copy to ToEvent
    '       New entry in CompEvents table
    '           All fields remain the same except for E_Code (generate by looking up in "Events in Full" query

  Set db = CurrentDb()
  Set Frs = db.OpenRecordset("CompEvents-With Event Type", dbOpenDynaset)   ' Create Recordset.
  Set Trs = db.OpenRecordset("CompEvents", dbOpenDynaset)   ' Create Recordset.
    
    'Stop
  
  If IsNull(Me!FmET_Code) Then
    MsgBox ("The FROM event cannot be empty.")
    Exit Sub
  End If
  
  If IsNull(Me!ToET_Code) Then
    MsgBox ("The TO event cannot be empty.")
    Exit Sub
  End If
  
  If Nz(FmAge) = "" Then Me!FmAge = "*"
  If Nz(Me!FmSex) = "" Then Me!FmSex = "*"
  If Nz(Me!FmHeat) = "" Then Me!FmHeat = "*"
  If Nz(Me!FmFinalLev) = "" Then Me!FmFinalLev = "*"
    
  Fcriteria = "[ET_Code]=" & FmET_Code & " AND [Age] like """ & FmAge & """ AND [Sex] like """ & FmSex & """"
  If Me!FmHeat <> "*" Then
    Fcriteria = Fcriteria & " AND [Heat]= " & FmHeat
  End If
  If Me!FmFinalLev <> "*" Then
    Fcriteria = Fcriteria & " AND [F_Lev]= " & FmFinalLev
  End If
  Frs.FindFirst Fcriteria

  X = 0
  FailedI = 0
  SuccessI = 0

  While Not Frs.NoMatch

    Retval = SysCmd(SYSCMD_SETSTATUS, "Processing competitor " & X)
    X = X + 1

    NewE_Code = DLookup("[E_Code]", "Events", "[ET_Code]=" & Me!ToET_Code & " AND [Sex]=""" & Frs![Sex] & """ AND [Age]=""" & Frs![Age] & """")
    If IsNull(NewE_Code) Then
      FailedI = FailedI + 1
    ElseIf IsNull(DLookup("[E_Code]", "Heats", "[E_Code] = " & NewE_Code & " AND [F_Lev]=" & Frs!F_Lev & " AND [Heat]=" & Frs!Heat)) Then
      FailedI = FailedI + 1
    Else
              
      Criteria = "[E_Code] = " & NewE_Code & " AND [F_Lev]=" & Frs!F_Lev & " AND [Heat]=" & Frs!Heat & " AND [PIN]=" & Frs!PIN
      Trs.FindFirst Criteria
      If Trs.NoMatch Then ' The competitor is not already enrolled in the event
        SuccessI = SuccessI + 1
          Trs.AddNew
          Trs!PIN = Frs!PIN
          Trs!E_Code = NewE_Code
          Trs!F_Lev = Frs!F_Lev
          Trs!Heat = Frs!Heat
          Trs!Place = Frs!Place
          Trs!Lane = Frs!Lane
          Trs!Result = Frs!Result
          Trs!nResult = Frs!nResult
          Trs!Memo = Frs!Memo
          Trs!Points = Frs!Points
          Trs.Update
      End If
      
    End If
    Frs.FindNext Fcriteria
  Wend
  
  Retval = SysCmd(SYSCMD_CLEARSTATUS)

  Frs.Close
  Trs.Close
  msg = "Copy complete.  " & SuccessI & " competitors copied."
  If FailedI > 0 Then
    msg = msg & " However " & FailedI & " of " & X & " heats could not be processed because the TO heat did not exist."
  End If
  Response = MsgBox(msg, vbInformation)

Exit_CopyCompetitors_Click:
    DoCmd.RunMacro "ClosePleaseWait"
    Set db = Nothing
    Exit Sub

Err_CopyCompetitors_Click:
    MsgBox Error$
    Resume Exit_CopyCompetitors_Click
    
End Sub

Private Sub FmAge_AfterUpdate()
    
    'Me!ToAge = Me!FmAge

End Sub

Private Sub FmAge_DblClick(Cancel As Integer)
  
  Me!FmAge = "*"
  
End Sub

Private Sub FmFinalLev_AfterUpdate()

    'Me!ToFinalLev = Me!FmFinalLev

End Sub

Private Sub FmFinalLev_DblClick(Cancel As Integer)
  Me!FmFinalLev = "*"
End Sub

Private Sub FmHeat_BeforeUpdate(Cancel As Integer)

    'Me!ToHeat = Me!FmHeat

End Sub

Private Sub FmHeat_DblClick(Cancel As Integer)
  Me!FmHeat = "*"
End Sub

Private Sub FmSex_AfterUpdate()

    'Me!ToSex = Me!FmSex

End Sub

Private Sub FmSex_DblClick(Cancel As Integer)
  Me!FmSex = "*"
End Sub
