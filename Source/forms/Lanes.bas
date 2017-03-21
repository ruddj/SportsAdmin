Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =6406
    ItemSuffix =41
    Left =2430
    Top =1035
    Right =9825
    Bottom =7380
    HelpContextId =130
    RecSrcDt = Begin
        0x23bb97290fcde140
    End
    Caption ="Default Lane Allocation"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
            Height =5280
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =3
                    Left =170
                    Top =377
                    Width =4087
                    Height =4787
                    Name ="Lanes SF"
                    SourceObject ="Form.Lanes SF"
                    EventProcPrefix ="Lanes_SF"

                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =453
                    Top =141
                    Width =810
                    Height =240
                    BackColor =-2147483633
                    Name ="Text36"
                    Caption ="Lane"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =1
                    Left =1256
                    Top =141
                    Width =1875
                    Height =240
                    BackColor =-2147483633
                    Name ="Text37"
                    Caption ="Team"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =4680
                    Top =4622
                    Width =1418
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Button28"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4705
                    Top =377
                    Width =1418
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="Order"
                    Caption ="Refresh Order"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Refreshes the list on the left so that it is in lane order."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4695
                    Top =1155
                    Width =1418
                    Height =532
                    FontSize =7
                    FontWeight =400
                    TabIndex =3
                    Name ="Update"
                    Caption ="Update Lanes for all Events"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    ControlTipText ="Put all competitors already in events in the lanes as specified in the list."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4680
                    Top =3960
                    Width =1418
                    Height =397
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    HelpContextId =40
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="MS Sans Serif"

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

Private Sub Button26_Click()
On Error GoTo Err_Button26_Click


    DoCmd.GoToRecord , , A_PREVIOUS

Exit_Button26_Click:
    Exit Sub

Err_Button26_Click:
    MsgBox Error$
    Resume Exit_Button26_Click
    
End Sub

Private Sub Button27_Click()
On Error GoTo Err_Button27_Click


    DoCmd.GoToRecord , , A_NEXT

Exit_Button27_Click:
    Exit Sub

Err_Button27_Click:
    MsgBox Error$
    Resume Exit_Button27_Click
    
End Sub

Private Sub Button28_Click()
On Error GoTo Err_Button28_Click


    DoCmd.Close

Exit_Button28_Click:
    Exit Sub

Err_Button28_Click:
    MsgBox Error$
    Resume Exit_Button28_Click
    
End Sub

Private Sub Button29_Click()
On Error GoTo Err_Button29_Click


    DoCmd.GoToRecord , , A_NEWREC

Exit_Button29_Click:
    Exit Sub

Err_Button29_Click:
    MsgBox Error$
    Resume Exit_Button29_Click
    
End Sub

Private Sub Button30_Click()
On Error GoTo Err_Button30_Click


    DoCmd.GoToRecord , , A_NEXT

Exit_Button30_Click:
    Exit Sub

Err_Button30_Click:
    MsgBox Error$
    Resume Exit_Button30_Click
    
End Sub

Private Sub Button31_Click()
On Error GoTo Err_Button31_Click


    DoCmd.GoToRecord , , A_PREVIOUS

Exit_Button31_Click:
    Exit Sub

Err_Button31_Click:
    MsgBox Error$
    Resume Exit_Button31_Click
    
End Sub

Private Sub Button39_Click()

    
End Sub

Private Sub Order_Click()

    Me![Lanes SF].Requery

End Sub

Private Sub Update_Click()

 On Error GoTo Update_Click_Error
 
 Response = MsgBox("This will update the lanes allocated to all competitors.  All old lane allocations will be removed and updated.  Do you wish to continue?", 36, "Update Lane Allocation")
 
 If Response = 6 Then
    Dim Criteria As String, Db As Database, rs As Recordset
    Q = "UPDATE DISTINCTROW CompEvents SET CompEvents.Lane = 0"
    DoCmd.SetWarnings False
    DoCmd.RunSQL Q
    DoCmd.SetWarnings True

    DoCmd.Hourglass True

    Set Db = DBEngine.Workspaces(0).Databases(0)
    Set rs = Db.OpenRecordset("Heats", dbOpenDynaset)   ' Create dynaset.

    msg = "Updating Lanes ... "
    ReturnValue = SysCmd(acSysCmdInitMeter, msg, DCount("[E_Code]", "Heats"))   ' Display message in status bar.
    x = 0

    If Not rs.BOF Then
    
        rs.MoveFirst
        
        Do Until rs.EOF
           Call Update_Lane_Assignments(rs!E_Code, rs!F_Lev, rs!Heat)
            x = x + 1
            ReturnValue = SysCmd(acSysCmdUpdateMeter, x)   ' Update meter.
    
            rs.MoveNext
        Loop
    
        rs.Close
    Else
        MsgBox ("There are no lane to be updated.")
    End If
 End If

Update_Click_Exit:
    DoCmd.Hourglass False
    ReturnValue = SysCmd(acSysCmdRemoveMeter)
    Exit Sub
    
Update_Click_Error:
    MsgBox (Error$)
    GoTo Update_Click_Exit
End Sub
