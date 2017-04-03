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
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =9924
    ItemSuffix =73
    Left =1560
    Top =1305
    Right =13755
    Bottom =10530
    RecSrcDt = Begin
        0xe130be95d6e5e140
    End
    RecordSource ="Misc-Utilities"
    Caption ="Utilities"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontFamily =2
            Width =283
            Height =283
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =6624
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =8280
                    Top =6090
                    Width =1584
                    Height =450
                    FontWeight =700
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
                    Left =60
                    Top =60
                    Width =9810
                    Height =5970
                    TabIndex =1
                    Name ="TabCtl"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =9870
                    LayoutCachedHeight =6030
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Page30"
                            Caption ="Substitutions"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =1
                                    Left =437
                                    Top =1146
                                    Width =3521
                                    Height =3775
                                    BorderColor =12632256
                                    Name ="Lane Subtitution SF"
                                    SourceObject ="Form.Lane Subtitution SF"
                                    EventProcPrefix ="Lane_Subtitution_SF"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =436
                                            Top =901
                                            Width =1695
                                            Height =240
                                            Name ="Text12"
                                            Caption ="Lane Substitiuion"
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =1
                                    Left =4200
                                    Top =4031
                                    Width =3135
                                    Height =892
                                    TabIndex =1
                                    BorderColor =12632256
                                    Name ="Embedded22"
                                    SourceObject ="Form.SexSub"

                                    LayoutCachedLeft =4200
                                    LayoutCachedTop =4031
                                    LayoutCachedWidth =7335
                                    LayoutCachedHeight =4923
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =4200
                                            Top =3795
                                            Width =1470
                                            Height =240
                                            Name ="Text23"
                                            Caption ="Gender Substitution"
                                            FontName ="MS Sans Serif"
                                            LayoutCachedLeft =4200
                                            LayoutCachedTop =3795
                                            LayoutCachedWidth =5670
                                            LayoutCachedHeight =4035
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =1
                                    Left =4182
                                    Top =1171
                                    Width =3521
                                    Height =2515
                                    TabIndex =2
                                    BorderColor =12632256
                                    Name ="Final Level SF"
                                    SourceObject ="Form.Final Level SF"
                                    EventProcPrefix ="Final_Level_SF"

                                    LayoutCachedLeft =4182
                                    LayoutCachedTop =1171
                                    LayoutCachedWidth =7703
                                    LayoutCachedHeight =3686
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =4181
                                            Top =926
                                            Width =1695
                                            Height =240
                                            Name ="Title"
                                            Caption ="Final Level Substitiuion"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Page33"
                            Caption ="Misc"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =470
                                    Top =738
                                    Width =1470
                                    Height =510
                                    Name ="Remove"
                                    Caption ="Delete All Competitors"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =470
                                    Top =2459
                                    Width =1470
                                    Height =510
                                    TabIndex =1
                                    Name ="Reset"
                                    Caption ="Reset All Events"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =2038
                                    Top =741
                                    Width =6000
                                    Height =630
                                    Name ="Label49"
                                    Caption ="This button will delete all competitors from your database.  This is useful if y"
                                        "ou have copied an exisiting carnival and are planning to import all the competit"
                                        "ors from a new text file."
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =340
                                    Top =623
                                    Width =7796
                                    Height =799
                                    Name ="Box50"
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =2038
                                    Top =2447
                                    Width =6000
                                    Height =1590
                                    Name ="Label51"
                                    Caption ="This will set all events and heats back to their original state.  The status of "
                                        "heats and final-levels change as your progress through a carnival (heats: incomp"
                                        "lete to complete, final-levels: future to active to completed to promoted).  Thi"
                                        "s action will set all completed heats back to not completed.  Also the first fin"
                                        "al-level in each event will become 'Active' with all future finals being set to "
                                        "future. \015\012\015\012Use this button if you have copied an exisiting carnival"
                                        " and want to take it back to its original state."
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =340
                                    Top =2333
                                    Width =7796
                                    Height =1759
                                    Name ="Box52"
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =470
                                    Top =1593
                                    Width =1470
                                    Height =510
                                    TabIndex =2
                                    Name ="ClearResults"
                                    Caption ="Clear all Results"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =2038
                                    Top =1596
                                    Width =6000
                                    Height =630
                                    Name ="Label54"
                                    Caption ="This button will remove all competitors from all events.  it does not remove the"
                                        " competitors from the database.  This is useful if you have copied an exisiting "
                                        "carnival and are not planning to import all the competitors from a new text file"
                                        "."
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =340
                                    Top =1478
                                    Width =7796
                                    Height =799
                                    Name ="Box55"
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =480
                                    Top =4305
                                    Width =1470
                                    Height =510
                                    TabIndex =3
                                    Name ="RecereateHeatsBut"
                                    Caption ="Recreate All Heats"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Tahoma"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Label
                                    OverlapFlags =223
                                    Left =2043
                                    Top =4299
                                    Width =6000
                                    Height =885
                                    Name ="Label71"
                                    Caption ="This will clear all competitors from all existing events and recreate the heats "
                                        "and finals as they are setup in the 'Quickly Setup Heats' form."
                                End
                                Begin Rectangle
                                    SpecialEffect =3
                                    BackStyle =0
                                    OverlapFlags =215
                                    Left =345
                                    Top =4185
                                    Width =7796
                                    Height =1054
                                    Name ="Box72"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =465
                            Width =9540
                            Height =5435
                            Name ="Page68"
                            Caption ="Titles"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5900
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =275
                                    Top =545
                                    Width =8010
                                    Height =5355
                                    Name ="Child69"
                                    SourceObject ="Form.Utilities 3"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Page34"
                            Caption ="HTML Settings"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    CanGrow = NotDefault
                                    OverlapFlags =247
                                    Left =356
                                    Top =866
                                    Width =7485
                                    Height =3945
                                    Name ="Utilities-HTML"
                                    SourceObject ="Form.Utilities-HTML"
                                    EventProcPrefix ="Utilities_HTML"

                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =6554
                                    Top =5240
                                    Width =1149
                                    Height =450
                                    TabIndex =1
                                    HelpContextId =540
                                    Name ="HelpHTML"
                                    Caption ="Help"
                                    OnClick ="Open Help"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Page37"
                            Caption ="Remove Empty Heats"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =341
                                    Top =626
                                    Width =7365
                                    Height =5190
                                    Name ="Utilities 2"
                                    SourceObject ="Form.Utilities 2"
                                    EventProcPrefix ="Utilities_2"

                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Copy Competitors"
                            EventProcPrefix ="Copy_Competitors"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =359
                                    Top =776
                                    Width =7065
                                    Height =4890
                                    Name ="CopyCompetitorsBetweenEvents"
                                    SourceObject ="Form.CopyCompetitorsBetweenEvents"

                                    LayoutCachedLeft =359
                                    LayoutCachedTop =776
                                    LayoutCachedWidth =7424
                                    LayoutCachedHeight =5666
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Page49"
                            Caption ="Backup"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =215
                                    Left =581
                                    Top =883
                                    BorderColor =12632256
                                    Name ="BackupToCarnivalPath"
                                    ControlSource ="BackupToCarnivalPath"
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =807
                                            Top =849
                                            Width =3300
                                            Height =240
                                            Name ="Label53"
                                            Caption ="Backup to the same folder as the carnival file"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    Left =1771
                                    Top =1189
                                    Width =4476
                                    TabIndex =1
                                    BorderColor =12632256
                                    Name ="BackupPath"
                                    ControlSource ="BackupPath"
                                    FontName ="Tahoma"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =581
                                            Top =1189
                                            Width =1140
                                            Height =240
                                            Name ="Label55"
                                            Caption ="Backup Folder:"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =6363
                                    Top =1133
                                    Width =576
                                    Height =351
                                    FontSize =10
                                    TabIndex =3
                                    Name ="BackupFolderBut"
                                    Caption ="..."
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
                                        0x000000000000000000000000
                                    End
                                    FontName ="System"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =195
                            Top =465
                            Width =9540
                            Height =5430
                            Name ="Page58"
                            Caption ="Report Popup"
                            LayoutCachedLeft =195
                            LayoutCachedTop =465
                            LayoutCachedWidth =9735
                            LayoutCachedHeight =5895
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =375
                                    Top =615
                                    Width =3795
                                    Height =1395
                                    Name ="ReportsPopUp-SF"
                                    SourceObject ="Form.ReportsPopUp-SF"
                                    EventProcPrefix ="ReportsPopUp_SF"

                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =850
                    Top =6066
                    Width =561
                    TabIndex =2
                    Name ="CurrentTab"
                    ControlSource ="CurrentTab"

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

Private Sub Button3_Click()

End Sub

Private Sub BackupFolderBut_Click()

Dim Result As String

On Error GoTo Err_BackupFolderBut_Click

    Dim n As Variant

    n = BrowseFolder("Locate web folder") 'Me!ctlCommonDialog.FileName
    If n <> "" Then
        Me![BackupPath] = ExtractDirectory(Trim(n))
    End If

Exit_BackupFolderBut_Click:
    Exit Sub

Err_BackupFolderBut_Click:
    MsgBox Error$
    Resume Exit_BackupFolderBut_Click
    


End Sub

Private Sub BackupToCarnivalPath_AfterUpdate()

  If Me!BackupToCarnivalPath Then
    Me!BackupPath.enabled = False
    Me!BackupFolderBut.enabled = False
  Else
    Me!BackupPath.enabled = True
    Me!BackupFolderBut.enabled = True
  End If

End Sub

Private Sub ClearResults_Click()

On Error GoTo ClearResults_Click_Err

  Dim Response As Variant
  
  Response = MsgBox("Are you sure you want to clear all results for this carnival?  This cannot be undone.", vbExclamation + vbYesNo + vbDefaultButton2)
  If Response = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "Delete * FROM [CompEvents]"
    DoCmd.SetWarnings True
  End If
  
ClearResults_Click_Exit:
  Exit Sub
  
ClearResults_Click_Err:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Close_Click()
On Error GoTo Close_Click_Err

    DoCmd.Close

Close_Click_Exit:
  Exit Sub
  
Close_Click_Err:
  MsgBox Err.Description, vbCritical
  
End Sub

Private Sub Edit_Lane_Click()


End Sub

Private Sub Edit_Lane_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me![Lane Subtitution SF].visible = True
    Me![Final Level SF].visible = False

End Sub

Private Sub Final_Level_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me![Lane Subtitution SF].visible = False
    Me![Final Level SF].visible = True

End Sub

Private Sub Form_Load()

  Me!TabCtl.Value = Me!CurrentTab
  Call BackupToCarnivalPath_AfterUpdate
  
End Sub

Private Sub HTMLbut_Click()

On Error GoTo Err_HTMLbut_Click

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "Utilities-HTML"
    DoCmd.OpenForm DocName, , , LinkCriteria

Exit_HTMLbut_Click:
    Exit Sub

Err_HTMLbut_Click:
    MsgBox Error$
    Resume Exit_HTMLbut_Click
    

End Sub

Private Sub MoreUtil_Click()
On Error GoTo Err_MoreUtil_Click

    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "Utilities 2"
    DoCmd.OpenForm DocName, , , LinkCriteria

Exit_MoreUtil_Click:
    Exit Sub

Err_MoreUtil_Click:
    MsgBox Error$
    Resume Exit_MoreUtil_Click
    
End Sub

Private Sub RecereateHeatsBut_Click()
On Error GoTo RecereateHeatsBut_Click_Err

  Response = MsgBox("Are you sure you want to re-create all heats?  (This will remove all competitors from all heats)", vbQuestion + vbDefaultButton2 + vbYesNo)
  If Response = vbNo Then Exit Sub

  Dim rs As Recordset
  Set rs = CurrentDb.OpenRecordset("EventType")
  
  If rs.BOF Then
    MsgBox "No events to process.", vbInformation
    Exit Sub
  End If
  
  Do Until rs.BOF Or rs.EOF
    If Not AutomaticallyCreateHeatsAndFinals(rs!ET_Code, , True) Then
      MsgBox "An error occurred recreating heats for: " & rs!ET_Des, vbExclamation
    End If
    rs.MoveNext
  Loop
  
  Response = MsgBox("Heats and finals have been created.", vbInformation)
  
  
RecereateHeatsBut_Click_Exit:
  Exit Sub
  
RecereateHeatsBut_Click_Err:
  MsgBox "An error occurred in RecereateHeatsBut_Click: " & Err.Description, vbCritical
  

End Sub

Private Sub Remove_Click()
On Error GoTo Remove_Click_Err

    Q = "DELETE DISTINCTROW Competitors.PIN FROM Competitors"

    r = MsgBox("Are you sure you want to remove all competitors from this carnival?", 36, "Remove all")
    If r = 6 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True
        Call TransferToCompetitorOrdered
        
    End If

Remove_Click_Exit:
  Exit Sub
  
Remove_Click_Err:
  MsgBox "An error occurred in Remove_Click: " & Err.Description, vbCritical
  
End Sub

Private Sub Reset_Click()
    On Error GoTo Reset_Click_Err
    
    ' Set the highest final level to active and the rest to future
    ' Set all compeitor points to 0

     r = MsgBox("Resetting all events will set the first heats to active and all others to future.  Are you sure you wish to reset all events?", 36, "Reset all events")
     If r = 6 Then
        
        PleaseWaitMsg = "Resetting all events ..."
        DoCmd.RunMacro "ShowPleaseWait"

    
        Dim Criteria As String, Db As Database, rs As Recordset
        
        Set Db = CurrentDb()
        Set rs = Db.OpenRecordset("SELECT * FROM Heats ORDER BY [E_CODE], [F_LEV] DESC ", dbOpenDynaset)   ' Create dynaset.
        
        rs.MoveFirst
        PrevE_Code = rs![E_Code]
        PrevF_Lev = rs![F_Lev]
        rs.Edit
        rs![Completed] = No
        rs![Status] = 1           ' 0=future; 1=active
        rs.Update
    
        rs.MoveNext
    
        CurrentlyActive = True
    
        Do Until rs.EOF
            
            rs.Edit
            rs![Completed] = No
            
    
            If rs![E_Code] = PrevE_Code Then
                
                If PrevF_Lev <> rs![F_Lev] Then
                    CurrentlyActive = False
                End If
                
                If CurrentlyActive Then
                    rs![Status] = 1           ' 0=future; 1=active
                Else
                    rs![Status] = 0
                End If
    
                PrevE_Code = rs![E_Code]
                PrevF_Lev = rs![F_Lev]
    
                rs.Update
                rs.MoveNext
    
            Else
                CurrentlyActive = True
                rs![Status] = 1
                PrevE_Code = rs![E_Code]
                PrevF_Lev = rs![F_Lev]
                rs.Update
                rs.MoveNext
            End If
    
        Loop
    
        rs.Close
    
        'q = "DELETE DISTINCTROW CompEvents.* FROM CompEvents"
        
        'DoCmd SetWarnings False
        'DoCmd RunSQL q
        'DoCmd SetWarnings False
        DoCmd.RunMacro "ClosePleaseWait"
     End If


Reset_Click_Exit:
    Set Db = Nothing
    Exit Sub

Reset_Click_Err:
    MsgBox (Error$)
    GoTo Reset_Click_Exit


End Sub

Private Sub TabCtl_Change()
  Me!CurrentTab = Me!TabCtl.Value
End Sub
