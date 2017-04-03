Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =20
    GridY =20
    Width =10368
    ItemSuffix =41
    Left =390
    Top =390
    Right =11430
    Bottom =8790
    HelpContextId =150
    RecSrcDt = Begin
        0x4e2adfb911cde140
    End
    RecordSource ="MiscellaneousLocal"
    Caption ="Import Carnival Disks"
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
        Begin CheckBox
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin Section
            CanGrow = NotDefault
            Height =7200
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    Left =144
                    Top =500
                    Width =3666
                    Height =285
                    BorderColor =12632256
                    Name ="FullFileName"
                    ControlSource ="ImportLocation"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"A:\\\""
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =146
                            Top =216
                            Width =2835
                            Height =285
                            Name ="Text1"
                            Caption ="Filename to Import (include full path):"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =5040
                    Top =264
                    Width =3385
                    Height =646
                    TabIndex =1
                    Name ="Format"
                    ControlSource ="ImportFormat"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5160
                            Top =144
                            Width =1016
                            Height =255
                            BackColor =-2147483633
                            Name ="Text3"
                            Caption ="Format Type"
                            FontName ="Tahoma"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =5184
                            Top =532
                            OptionValue =1
                            BorderColor =12632256
                            Name ="Field5"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    Left =5455
                                    Top =504
                                    Width =1005
                                    Height =240
                                    Name ="Text6"
                                    Caption ="Plain Text"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =6765
                            Top =538
                            OptionValue =2
                            BorderColor =12632256
                            Name ="Field7"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    Left =7050
                                    Top =510
                                    Width =1230
                                    Height =240
                                    Name ="Text8"
                                    Caption ="Plain Text (CSV)"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8856
                    Top =6336
                    Width =1374
                    Height =510
                    FontSize =8
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
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8856
                    Top =864
                    Width =1376
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =8404992
                    Name ="View"
                    Caption ="Show Text File"
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
                    Top =3816
                    Width =1376
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =4
                    ForeColor =32768
                    Name ="ImportData"
                    Caption ="Import All Displayed Data"
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
                    TextAlign =2
                    Left =432
                    Top =1080
                    Width =660
                    Height =210
                    Name ="Text14"
                    Caption ="House"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1152
                    Top =1080
                    Width =675
                    Height =225
                    Name ="Text15"
                    Caption ="Sex"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =1933
                    Top =1080
                    Width =555
                    Height =225
                    Name ="Text16"
                    Caption ="Age"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =2613
                    Top =1080
                    Width =1920
                    Height =225
                    Name ="Text17"
                    Caption ="Event"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4608
                    Top =1080
                    Width =495
                    Height =225
                    Name ="Text18"
                    Caption ="Heat"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =5112
                    Top =1080
                    Width =285
                    Height =225
                    Name ="Text19"
                    Caption ="#"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =5328
                    Top =1080
                    Width =1155
                    Height =225
                    Name ="Text20"
                    Caption ="Given Name"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =6575
                    Top =1080
                    Width =1230
                    Height =225
                    Name ="Text21"
                    Caption ="Surname"
                    FontName ="Tahoma"
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =144
                    Top =1368
                    Width =8353
                    Height =5581
                    TabIndex =5
                    Name ="I_Data"
                    SourceObject ="Form.Import Data SF"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8856
                    Top =216
                    Width =1376
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    Name ="Button24"
                    Caption ="Clear Temporary Data"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =223
                    Left =8640
                    Top =2512
                    Width =696
                    TabIndex =7
                    Name ="Field25"
                    ControlSource ="OpenAge"
                    DefaultValue ="\"A:\\\""
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =8641
                            Top =2232
                            Width =840
                            Height =285
                            Name ="Text26"
                            Caption ="Open Age:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =8640
                    Top =2882
                    TabIndex =8
                    BorderColor =12632256
                    Name ="SortByName"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =85
                            Left =8898
                            Top =2880
                            Width =1035
                            Height =240
                            Name ="Text29"
                            Caption ="Sort by Name"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3911
                    Top =455
                    Width =576
                    Height =366
                    FontSize =8
                    FontWeight =400
                    TabIndex =9
                    Name ="Locate"
                    Caption ="Locate"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadad00000000000adada ,
                        0x003333333330adad0b03333333330ada0fb03333333330ad0bfb03333333330a ,
                        0x0fbfb000000000000bfbfbfbfb0adada0fbfbfbfbf0dadad0bfb0000000adada ,
                        0xa000adadadad000ddadadadadadad00aadadadad0dad0d0ddadadadad000dada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
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

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Line
                    OverlapFlags =119
                    Left =8640
                    Width =0
                    Height =7200
                    Name ="Line38"
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

Private Sub Button24_Click()

    Mesg = "This will remove all the data shown on the previous form.  This will not effect the Carnival Database.  Do you wish to contnue?"
    Response = MsgBox(Mesg, 20)
    If Response = 6 Then
      CurrentDb.Execute "DELETE DISTINCTROW ImportData.* FROM ImportData"
      Me.I_Data.Requery
    End If

End Sub

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

Private Sub FullFileName_AfterUpdate()

  If Right(Me.FullFileName, 3) = "CSV" Then
    Me.Format = 2
  ElseIf Right(Me.FullFileName, 3) = "TXT" Then
    Me.Format = 1
  End If
  
End Sub

Private Sub ImportData_Click()
On Error GoTo Err_ImportData_Click

  Response = MsgBox("Push yes to continue with this import.", vbYesNo + vbQuestion)
  If Response = vbNo Then Exit Sub
  
  PleaseWaitMsg = "Importing competitors from carnival disks.  Please wait ..."
  DoCmd.RunMacro "ShowPleaseWait"

  DoCmd.Hourglass True
  DoCmd.RunCommand acCmdSaveRecord
  
  Dim Criteria As String, Db As Database, rs As Recordset, Q
  Dim Crs As Recordset, ITRS As Recordset, CErs As Recordset, NewAge As Variant

  Dim NamesIncomplete  As Integer, H_CodeIncomplete As Integer, msg As Variant
  Dim ShowCompetitorAlreadyEnrolledMessage As Boolean
  
  Set Db = CurrentDb()
  Set ITRS = Db.OpenRecordset("SELECT * FROM ImportData ORDER BY [Age]", dbOpenDynaset)   ' Create dynaset.
  Set Crs = Db.OpenRecordset("SELECT * FROM Competitors ORDER BY [Age] DESC", dbOpenDynaset)   ' Create dynaset.
  Set CErs = Db.OpenRecordset("CompEvents", dbOpenDynaset)   ' Create dynaset.
  
  ShowCompetitorAlreadyEnrolledMessage = True
  NamesIncomplete = 0
  H_CodeIncomplete = 0
  
  If ITRS.BOF Then
    MsgBox ("There is no data to import.")
  Else
    ITRS.MoveFirst
    Continue = True
    
    msg = "Processing competitor ... "
    'ReturnValue = SysCmd(acSysCmdInitMeter, Msg, DCount("[HE_Code]", "ImportData"))   ' Display message in status bar.
    ReturnValue = SysCmd(acSysCmdSetStatus, msg)
    X = 0

    While Not ITRS.EOF And Continue
      X = X + 1
      'ReturnValue = SysCmd(acSysCmdUpdateMeter, X)   ' Update meter.
      ReturnValue = ReturnValue = SysCmd(acSysCmdSetStatus, msg & X)
      ActualSex = DetermineSex(ITRS!Sex)
      
      If IsNull(ITRS!G_name) Or IsNull(ITRS!S_name) Then
        NamesIncomplete = NamesIncomplete + 1
      
      ElseIf IsNull(ITRS!H_Code) Then
        H_CodeIncomplete = H_CodeIncomplete + 1
      
      ElseIf IsNull(ITRS!HE_Code) Then
        H_CodeIncomplete = H_CodeIncomplete + 1
      
      ElseIf IsNull(DLookup("[HE_Code]", "Heats", "[HE_Code]=" & ITRS!HE_Code)) Then
        H_CodeIncomplete = H_CodeIncomplete + 1
      
      ElseIf ActualSex = False Then
        NamesIncomplete = NamesIncomplete + 1
        
      ElseIf CompetitorEnrolled(ITRS!G_name, ITRS!S_name) Then

        ActualSex = DetermineSex(ITRS!Sex)
        If ITRS!Age = "OPEN" Then ' competitor could be any age so look for Oldest
                                  'competior with name minus age
          Criteria = "[Gname]=""" & ITRS!G_name & """ AND [Surname]=""" & ITRS!S_name & """ AND [H_Code]=""" & ITRS!H_Code & """" & " AND [Sex]=""" & ActualSex & """"
          NewAge = 99
        Else
          NewAge = DetermineAge(ITRS!Age)
          Criteria = "[Gname]=""" & ITRS!G_name & """ AND [Surname]=""" & ITRS!S_name & """ AND [Age] =" & NewAge & " AND [H_Code]=""" & ITRS!H_Code & """" & " AND [Sex]=""" & ActualSex & """"
        End If
        Crs.FindFirst Criteria
        
        If Crs.NoMatch Then
            Crs.AddNew
            
            Crs!Include = True
            Crs!Gname = ITRS!G_name
            Crs!Surname = ITRS!S_name
            Crs!Sex = ActualSex
            Crs!H_Code = ITRS!H_Code
            Crs!H_ID = DetermineH_ID(ITRS!H_Code)
            Crs!Age = NewAge
            Crs!DOB = DetermineDOB(ITRS!Age)
            Crs!TotPts = 0

            Crs.Update
            
            Crs.MoveLast
            CompPIN = Crs!PIN

        Else
            CompPIN = Crs!PIN
        End If

        On Error GoTo ErrorAddingCompetitorToHeat
        
        F_Lev = DLookup("[F_Lev]", "Heats", "[HE_Code]=" & ITRS!HE_Code)
        E_Code = DLookup("[E_Code]", "Heats", "[HE_Code]=" & ITRS!HE_Code)
        Heat = DetermineHeat(ITRS!Heat)

        Criteria = "[PIN]=" & CompPIN & " AND [E_Code]=" & E_Code & " AND [F_Lev]= " & F_Lev & " AND [Heat]=" & Heat
        
        CErs.FindFirst Criteria

        If CErs.NoMatch Then
            CErs.AddNew

            CErs!PIN = CompPIN
            CErs!E_Code = E_Code
            CErs!Place = 0
            CErs!F_Lev = F_Lev
            CErs!Heat = Heat
            CErs!Lane = 0
            CErs!Result = 0
            CErs!nResult = 0
            
            CErs.Update
            On Error GoTo Err_ImportData_Click
            Call Update_Lane_Assignments(E_Code, F_Lev, Heat)

        ElseIf ShowCompetitorAlreadyEnrolledMessage Then
            ermsg = "The competitor " & ITRS!G_name & " " & ITRS!S_name & " from " & ITRS!H_Code & " is already enrolled in age " & ITRS!Age & " " & ITRS!Sex & " " & ITRS!ET_Des & " event.  You can only enrol a competitor once in an event. "
            ermsg = ermsg & "Push Yes to continue seeing these error messages.  Push No to continue importing but without the error messages.  Push Cancel to cancel the import."
            Response = MsgBox(ermsg, vbInformation + vbYesNoCancel)
            
            If Response = vbCancel Then
              Continue = False
            ElseIf Response = vbNo Then
              ShowCompetitorAlreadyEnrolledMessage = False
            End If

        End If
        
ResumeAddingCompetitorToHeat:

      End If

      ITRS.MoveNext
       
    Wend
    
    Call TransferToCompetitorOrdered
    
    'ReturnValue = SysCmd(acSysCmdRemoveMeter)
    ReturnValue = SysCmd(acSysCmdClearStatus)
    
    
    DoCmd.Beep
    msg = ""
    If H_CodeIncomplete > 0 Then
      msg = H_CodeIncomplete & " entries had improper EVENT data. "
    End If
    If NamesIncomplete > 0 Then
      msg = msg & NamesIncomplete & " entries had INCOMPLETE NAMES. "
    End If
    
    msg = msg & "The import is now complete."
    
    MsgBox msg, vbInformation


  End If
  
Exit_ImportData_Click:
    On Error Resume Next
    Crs.Close
    CErs.Close
    
    Set Db = Nothing
    DoCmd.Hourglass False
    DoCmd.RunMacro "ClosePleaseWait"

    Exit Sub

Err_ImportData_Click:
    'Q = "An error has been encountered during the import process:" & vbCr & vbLf & vbCr & vbLf
    'Q = Q & Err.Description
    Q = "There text file(s) that you are attempting to import appear to have corrupt data.  Please check that they are correct." & vbLf & vbCr
    Q = Q & " ERROR: " & Error$
    MsgBox Q, vbCritical
    Resume Exit_ImportData_Click
    
    
ErrorAddingCompetitorToHeat:
  Q = "A problem has occured adding " & ITRS!G_name & " " & ITRS!S_name & " to the race " & ITRS!Sex & " " & ITRS!Age & " " & ITRS!ET_Des & " " & ITRS!Heat
  Q = Q & ". This person will not be added. It will probably be easiest to add this competitor manually once the import is complete."
  Q = Q & "Push Yes to continue importing.  Push No cancel importing."
  Response = MsgBox(Q, vbInformation + vbYesNo)
  If Response = vbNo Then Continue = False
  
  GoTo ResumeAddingCompetitorToHeat
  
End Sub

Private Sub Locate_Click()

    Dim n As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    
    If Me.Format = 1 Then
        strFilter = ahtAddFilterItem(strFilter, "Team Text File (*.txt)", "*.txt")
    Else
        strFilter = ahtAddFilterItem(strFilter, "Team Excel File (*.xls)", "*.xls")
    End If
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    n = ahtCommonFileOpenSave(InitialDir:="", _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate Team Text File")
    If n <> "" Then
        Me![FullFileName] = Trim(n)
        Call FullFileName_AfterUpdate
    End If


    
End Sub

Private Sub SortByName_AfterUpdate()

    If Me![SortByName] Then
        Q = "SELECT DISTINCTROW ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age, ImportData.Sex, ImportData.HE_Code, ImportData.ET_Des, ImportData.Heat, ImportData.Competitor, ImportData.Memo "
        Q = Q & "FROM ImportData ORDER BY ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age"
        Me![I_Data].Form.RecordSource = Q
    Else
        Q = "SELECT DISTINCTROW ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age, ImportData.Sex, ImportData.HE_Code, ImportData.ET_Des, ImportData.Heat, ImportData.Competitor, ImportData.Memo "
        Q = Q & "FROM ImportData " ' ORDER BY ImportData.S_Name, ImportData.G_Name, ImportData.H_Code, ImportData.Age"
        Me![I_Data].Form.RecordSource = Q
    End If
    
    Me![I_Data].Requery


End Sub

Private Sub View_Click()

On Error GoTo Err_Import_Click

    FileN = Me![FullFileName]
    Mesg = "This will retrieve the data found in the file " & FileN & ".  It will be appended to the temporary table displayed on the previous form for you to view.  If you do not wish to append the data, delete the contents of the temporary table by pushing the 'Clear Temporary Data' button.  Do you wish to continue?"
    Response = MsgBox(Mesg, vbQuestion + vbYesNo)
    If Response = vbYes Then

        If Me.Format = 1 Then
            DoCmd.TransferText acImportDelim, "Create Carnival Disks", "ImportData", FullFileName
        ElseIf Me.Format = 2 Then
            DoCmd.TransferText acImportDelim, "Create Carnival Disks", "ImportData", FullFileName, True
        Else
          
          Dim tdf As TableDef
          Dim Db As Database
          Dim rs As Recordset
            
        End If

        Me.I_Data.Requery
    End If


Exit_Import_Click:
    Exit Sub

Err_Import_Click:
    MsgBox "An error has occured importing the carnival disk.  Check that the file is in the correct format (all the commas in the correct place and none missing etc.).  The specific error message is: " & Error$, vbInformation
    Resume Exit_Import_Click
    

End Sub
