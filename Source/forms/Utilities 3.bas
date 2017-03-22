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
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =7961
    ItemSuffix =56
    Left =-18465
    Top =5265
    Right =-10710
    Bottom =10365
    RecSrcDt = Begin
        0x6bd443042dc7e140
    End
    RecordSource ="Miscellaneous"
    Caption ="Utilities"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
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
        Begin Section
            CanGrow = NotDefault
            Height =5284
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =215
                    Left =1251
                    Top =850
                    Width =666
                    BorderColor =12632256
                    Name ="Field18"
                    ControlSource ="OpenAge"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =101
                            Top =848
                            Width =1170
                            Height =240
                            Name ="Text19"
                            Caption ="Open Age:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =295
                    Top =1720
                    Width =1470
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    ForeColor =128
                    Name ="Remove"
                    Caption ="Delete All Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =295
                    Top =3441
                    Width =1470
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    ForeColor =128
                    Name ="Reset"
                    Caption ="Reset All Events"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1255
                    Top =113
                    Width =4821
                    TabIndex =3
                    BorderColor =12632256
                    Name ="CarnivalTitle"
                    ControlSource ="CarnivalTitle"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =113
                            Width =1170
                            Height =240
                            Name ="Text15"
                            Caption ="Report Header:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1255
                    Top =473
                    Width =4821
                    TabIndex =4
                    BorderColor =12632256
                    Name ="CarnivalFooter"
                    ControlSource ="CarnivalFooter"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =473
                            Width =1170
                            Height =240
                            Name ="Text17"
                            Caption ="Report Footer:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =1251
                    Top =1212
                    Width =2316
                    TabIndex =5
                    BorderColor =12632256
                    Name ="Text47"
                    ControlSource ="ImportLocation"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextAlign =1
                            Left =101
                            Top =1210
                            Width =1170
                            Height =240
                            Name ="Label48"
                            Caption ="Password:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1863
                    Top =1723
                    Width =6000
                    Height =630
                    Name ="Label49"
                    Caption ="This button will delete all competitors from your database.  This is useful if y"
                        "ou have copied an exisiting carnival and are planning to import all the competit"
                        "ors from a new text file."
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =215
                    Left =165
                    Top =1605
                    Width =7796
                    Height =799
                    Name ="Box50"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1863
                    Top =3429
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
                    Visible = NotDefault
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =215
                    Left =165
                    Top =3315
                    Width =7796
                    Height =1759
                    Name ="Box52"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontFamily =34
                    Left =295
                    Top =2575
                    Width =1470
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    ForeColor =128
                    Name ="ClearResults"
                    Caption ="Clear all Results"
                    OnClick ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =1863
                    Top =2578
                    Width =6000
                    Height =630
                    Name ="Label54"
                    Caption ="This button will remove all competitors from all events.  it does not remove the"
                        " competitors from the database.  This is useful if you have copied an exisiting "
                        "carnival and are not planning to import all the competitors from a new text file"
                        "."
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =215
                    Left =165
                    Top =2460
                    Width =7796
                    Height =799
                    Name ="Box55"
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


Private Sub Remove_Click()

    Q = "DELETE DISTINCTROW Competitors.PIN FROM Competitors"

    r = MsgBox("Are you sure you want to remove all competitors from this carnival?", 36, "Remove all")
    If r = 6 Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True
        Call TransferToCompetitorOrdered
        
    End If

    
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
        
        Set Db = DBEngine.Workspaces(0).Databases(0)
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
    Exit Sub

Reset_Click_Err:
    MsgBox (Error$)
    GoTo Reset_Click_Exit

End Sub
