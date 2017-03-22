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
    GridX =20
    GridY =20
    Width =10370
    ItemSuffix =48
    Left =645
    Top =1290
    Right =11010
    Bottom =8070
    HelpContextId =120
    RecSrcDt = Begin
        0xbc1d08fbafdce140
    End
    RecordSource ="Misc-ImportCompetitors"
    Caption ="Import Competitors"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
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
            Height =6789
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    Left =257
                    Top =428
                    Width =6576
                    TabIndex =1
                    BorderColor =12632256
                    Name ="FullFileName"
                    ControlSource ="ImportCompetitors"
                    DefaultValue ="\"A:\\\""
                    ControlTipText ="Enter the location of the competitors text file to import."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =259
                            Top =144
                            Width =3570
                            Height =285
                            FontWeight =700
                            Name ="Text1"
                            Caption ="Filename to Import (include full path):"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =8846
                    Top =6130
                    Width =1464
                    Height =510
                    FontSize =8
                    TabIndex =2
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8813
                    Top =3316
                    Width =1464
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    ForeColor =8404992
                    Name ="View"
                    Caption ="View Text File"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="View the text file you want to import in the list on the left."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8820
                    Top =4200
                    Width =1464
                    Height =510
                    FontSize =8
                    TabIndex =4
                    ForeColor =32768
                    Name ="ImportData"
                    Caption ="Import Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Import all competitors in the list on the left into the carnival."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4025
                    Top =870
                    Width =735
                    Height =210
                    FontWeight =700
                    Name ="Text14"
                    Caption ="House"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4980
                    Top =870
                    Width =645
                    Height =225
                    FontWeight =700
                    Name ="Text15"
                    Caption ="Sex"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5782
                    Top =870
                    Width =525
                    Height =225
                    FontWeight =700
                    Name ="Text16"
                    Caption ="Age"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =510
                    Top =870
                    Width =1230
                    Height =225
                    FontWeight =700
                    Name ="Text20"
                    Caption ="Given Name"
                    FontName ="Arial"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2268
                    Top =870
                    Width =1305
                    Height =225
                    FontWeight =700
                    Name ="Text21"
                    Caption ="Surname"
                    FontName ="Arial"
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =90
                    Top =1125
                    Width =8428
                    Height =5506
                    Name ="I_Data"
                    SourceObject ="Form.Import Competitors SF"

                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8813
                    Top =2636
                    Width =1464
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =5
                    Name ="btnClearTemp"
                    Caption ="Clear Temporary Data"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Clears all the competitors shown in the list on the left."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6519
                    Top =870
                    Width =630
                    Height =225
                    FontWeight =700
                    Name ="Text30"
                    Caption ="DOB"
                    FontName ="Arial"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6888
                    Top =396
                    Width =576
                    Height =366
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
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
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Click to locate the competitors text file to import."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =8846
                    Top =5450
                    Width =1464
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    HelpContextId =120
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =8610
                    Width =0
                    Height =6789
                    Name ="Line37"
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    Left =8938
                    Top =453
                    Width =1251
                    TabIndex =8
                    BorderColor =12632256
                    Name ="Text39"
                    ControlSource ="=Format(Now(),\"dd/mm/yyyy\")"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8945
                            Top =198
                            Width =1365
                            Height =240
                            Name ="Label40"
                            Caption ="Current Date:"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =975
                    Left =8938
                    Top =1474
                    Width =1176
                    TabIndex =9
                    BorderColor =12632256
                    Name ="CutMonth"
                    ControlSource ="MthAgeYearEnds"
                    RowSourceType ="Value List"
                    RowSource ="1;\"January\";2;\"February\";3;\"March\";4;\"April\";5;\"May\";6;\"June\";7;\"Ju"
                        "ly\";8;\"August\";9;\"September\";10;\"October\";11;\"November\";12;\"December\""
                    ColumnWidths ="0;975"
                    StatusBarText ="Choose the month when the competitor is moved into the next age division."
                    ControlTipText ="Choose the month when the competitor is moved into the next age division."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8951
                            Top =1190
                            Width =795
                            Height =240
                            Name ="Month_Label"
                            Caption ="Month"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    ListWidth =975
                    Left =8966
                    Top =2041
                    Width =1176
                    TabIndex =10
                    BorderColor =12632256
                    Name ="CutDay"
                    ControlSource ="DayAgeYearEnds"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30"
                        ";31"
                    ColumnWidths ="975"
                    StatusBarText ="Choose the day of that month when the competitor is moved into the next age divi"
                        "sion."
                    ControlTipText ="Choose the day of that month when the competitor is moved into the next age divi"
                        "sion."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =8979
                            Top =1757
                            Width =795
                            Height =240
                            Name ="Label44"
                            Caption ="Day"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =8825
                    Top =1015
                    Width =1417
                    Height =1366
                    Name ="Box45"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =8888
                    Top =930
                    Width =1290
                    Height =210
                    BackColor =-2147483633
                    Name ="Label46"
                    Caption ="Age Cut-Off Date"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7344
                    Top =876
                    Width =630
                    Height =225
                    FontWeight =700
                    Name ="Label47"
                    Caption ="PIN"
                    FontName ="Arial"
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
'Option Explicit


Private Sub Button27_Click()
On Error GoTo Err_Button27_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Button27_Click:
    Exit Sub

Err_Button27_Click:
    MsgBox Error$
    Resume Exit_Button27_Click
    
End Sub

Private Sub btnClearTemp_Click()
    Mesg = "This will clear any imported data from the temporary table.  This action does not effect the Carnival Database itself.  Do you wish to contnue?"
    Response = MsgBox(Mesg, 20)
    If Response = 6 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE DISTINCTROW [Import Competitors].* FROM [Import Competitors]"
        DoCmd.SetWarnings True
        [I_Data].Requery
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

Private Function CompetitorEnrolled(Gname, Sname)

    If IsNull(Gname) And IsNull(Sname) Then
        CompetitorEnrolled = False
    Else
        CompetitorEnrolled = True
    End If

End Function

Private Sub ImportData_Click()
On Error GoTo Err_ImportData_Click

  DoCmd.RunCommand acCmdSaveRecord
  
  Dim Criteria As String, Db As Database, rs As Recordset
  Dim Crs As Recordset, ITRS As Recordset, CErs As Recordset
  Dim Hrs As Recordset, H_ID As Variant

  Set Db = DBEngine.Workspaces(0).Databases(0)
  Set ITRS = Db.OpenRecordset("Import Competitors", dbOpenDynaset)   ' Create dynaset.
  Set Crs = OpenForSeek("Competitors")
  
  On Error Resume Next
  Crs.index = "Name&House"
  
  If Err.Number <> 0 Then Set Crs = Db.OpenRecordset("Competitors", dbOpenDynaset)   ' Create dynaset.
  
  On Error GoTo Err_ImportData_Click
  
  Set Hrs = Db.OpenRecordset("House", dbOpenDynaset)
    
  If ITRS.BOF Then
    Response = MsgBox("There are no competitors in the list to import.  First 'View the Text File' for correctness then push the import button.", vbInformation)
  Else
    ITRS.MoveFirst
    Continue = True
    PleaseWaitMsg = "Importing competitors.  Please wait ..."
    DoCmd.RunMacro "ShowPleaseWait"
    msg = "Importing Competitors ... "
    ReturnValue = SysCmd(acSysCmdInitMeter, msg, DCount("[Gname]", "Import Competitors"))   ' Display message in status bar.
    
    While Not ITRS.EOF And Continue
      x = x + 1
      ReturnValue = SysCmd(acSysCmdUpdateMeter, x)   ' Update meter.

      Cname = ITRS!Gname & " " & ITRS!Sname & " (" & UCase(ITRS!H_Code) & ")"
      Response = 0
    
      
      If IsNull(ITRS!Gname) Then
        Response = MsgBox("The GIVEN NAME for competitor " & Cname & " is not complete.  Please fix the entry and import the file again.  Do you wish to continue?", 20, "Import Competitors")
      ElseIf IsNull(ITRS!Sname) Then
        Response = MsgBox("The SURNAME for competitor " & Cname & " is not complete.  Please fix the entry and import the file again.  Do you wish to continue?", 20, "Import Competitors")
      ElseIf IsNull(ITRS!Sex) Then
        Response = MsgBox("The SEX FIELD for competitor " & Cname & " is not complete.  Please fix the entry and import the file again.  Do you wish to continue?", 20, "Import Competitors")
      ElseIf IsNull(ITRS!Age) And IsNull(ITRS!DOB) Then
        Response = MsgBox("Both the AGE and DOB fields are empty for competitor " & Cname & ".  Please fix the entry and import the file again.  Do you wish to continue?", 20, "Import Competitors")
      ElseIf IsNull(ITRS!H_Code) Then
        Response = MsgBox("The TEAM NAME for competitor " & Cname & " is empty.  Please fix the entry and import the file again.  Do you wish to continue?", 20, "Import Competitors")
      Else
         Hcode = UCase(ITRS!H_Code)
         Hrs.FindFirst "[H_Code]=""" & Hcode & """"
         'If IsNull(DLookup("[H_Code]", "House", "[H_Code]=""" & UCase(ITRS!H_Code) & """")) Then
         If Hrs.NoMatch Then
            
            Response = MsgBox("The team " & Hcode & " does not exist.  The Sports Administrator is now creating it.", vbInformation)
            
            Hrs.AddNew
            Hrs!H_Code = Hcode
            Hrs!H_NAme = Hcode
            Hrs!Include = True
            Hrs!CompPool = 0
            Hrs.Update
            Hrs.Bookmark = Hrs.LastModified ' Move to new record
         End If
         
         H_ID = Hrs!H_ID
        
         ActualSex = DetermineSex(ITRS!Sex)

         If IsNull(ITRS!Age) Then
          ActualAge = DetermineAge_ImportCompetitors(ITRS!DOB, Me!CutDay, Me!CutMonth)
         Else
            ActualAge = ITRS!Age
         End If
        
         If IsNull(ITRS!DOB) Then
            ActualDOB = DetermineDOB(ActualAge)
         ElseIf Not IsDate(ITRS!DOB) Then
            ActualDOB = DetermineDOB(ActualAge)
         Else
            ActualDOB = ITRS!DOB
         End If

          On Error Resume Next
          Crs.Seek "=", ITRS!Sname, ITRS!Gname, ActualAge, Hcode, ActualSex
          If Err.Number <> 0 Then
            On Error GoTo Err_ImportData_Click
            Criteria = "[Gname]=""" & ITRS!Gname & """ AND [Surname]=""" & ITRS!Sname & """ AND [Age] =" & ActualAge & " AND [H_Code]=""" & Hcode & """" & " AND [Sex]=""" & ActualSex & """"
            Crs.FindFirst Criteria
          End If
        
         If Crs.NoMatch Then
            Crs.AddNew
            
            Crs!Include = True
            Crs!Gname = ITRS!Gname
            Crs!Surname = ITRS!Sname
            Crs!Sex = ActualSex
            Crs!H_Code = Hcode
            Crs!H_ID = H_ID ' DetermineH_ID(Hcode)
            Crs!Age = ActualAge
            Crs!DOB = ActualDOB
            Crs!ID = Nz(ITRS!PIN) ' The new id field gets assigned the school
            Crs!TotPts = 0

            Crs.Update
            
            'crs.MoveLast
            'CompPIN = crs!PIN

         End If

       End If

       ITRS.MoveNext
      
       If Response = 7 Then
         Continue = False
       End If
       
     Wend
     Call TransferToCompetitorOrdered
     ReturnValue = SysCmd(acSysCmdRemoveMeter)
     DoCmd.Beep
     Response = MsgBox("The import is now complete.", vbInformation)
     
     Crs.Close
  End If
  
Exit_ImportData_Click:
    DoCmd.RunMacro "ClosePleaseWait"
    Exit Sub

Err_ImportData_Click:
    MsgBox Error$
    Resume Exit_ImportData_Click
    
End Sub

Private Sub Locate_Click()

On Error GoTo Locate_Click_Err
    
    Dim n As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    strFilter = ahtAddFilterItem(strFilter, "CSV, Tab, Web Delmited Files (*.csv;*.txt)", "*.csv;*.txt")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    n = ahtCommonFileOpenSave(InitialDir:="", _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate Competitor File")

    If n <> "" Then
        Me![FullFileName] = Trim(n)
    End If

Locate_Click_Exit:
  Exit Sub
    
Locate_Click_Err:
  MsgBox (Error$)
  GoTo Locate_Click_Exit
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
    
    Me![I_Data].Form.Requery


End Sub

Private Sub View_Click()

On Error GoTo Err_Import_Click

  FileN = Me![FullFileName]
  If Dir(FileN) <> "" Then
    If NoFormRecords(Me!I_Data.Form.RecordsetClone) Then ' No records in temporary table so don't display message
      Response = vbYes
    Else
      Mesg = "This will retrieve the data found in the file " & FileN & ".  It will be appended to the list of competitors displayed on the previous form for you to view.  If you do not wish to append the data, delete the contents of the temporary table by pushing the 'Clear Temporary Data' button.  Do you wish to continue?"
      Response = MsgBox(Mesg, vbYesNo + vbInformation)
    End If
    If Response = vbYes Then
        DoCmd.TransferText acImportDelim, "Import Competitors", "Import Competitors", FullFileName
        [I_Data].Requery ' Refresh Subform
    End If
  Else
    Response = MsgBox("The text file cannot be found in the specified position. Make sure you have entered the location of this file correctly.", vbCritical)
  End If

Exit_Import_Click:
    Exit Sub

Err_Import_Click:
    MsgBox Error$
    Resume Exit_Import_Click
    

End Sub
