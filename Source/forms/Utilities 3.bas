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
    ItemSuffix =57
    Left =-20160
    Top =7785
    Right =-10890
    Bottom =14730
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
                    FontName ="Tahoma"

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
                    FontName ="Tahoma"

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
                    FontName ="Tahoma"

                    LayoutCachedLeft =295
                    LayoutCachedTop =3441
                    LayoutCachedWidth =1765
                    LayoutCachedHeight =3951
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
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =113
                            Width =1170
                            Height =240
                            Name ="Text15"
                            Caption ="Report Header:"
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
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =56
                            Top =473
                            Width =1170
                            Height =240
                            Name ="Text17"
                            Caption ="Report Footer:"
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
                    FontName ="Tahoma"

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
                    OverlapFlags =223
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
                    FontName ="Tahoma"

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
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =295
                    Top =4140
                    Width =1470
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    ForeColor =128
                    Name ="ResetPoints"
                    Caption ="Reset Points"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Reset Extra Points from Teams"

                    LayoutCachedLeft =295
                    LayoutCachedTop =4140
                    LayoutCachedWidth =1765
                    LayoutCachedHeight =4650
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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

    r = MsgBox("Are you sure you want to remove all competitors from this carnival?", vbYesNo + vbQuestion, "Remove all")
    If r = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True
        Call TransferToCompetitorOrdered
        
    End If

    
End Sub

Private Sub Reset_Click()
On Error GoTo Reset_Click_Err
    Dim r As Boolean
    
    ' Set the highest final level to active and the rest to future
    ' Set all compeitor points to 0

     r = MsgBox("Resetting all events will set the first heats to active and all others to future.  Are you sure you wish to reset all events?", vbYesNo + vbQuestion, "Reset all events")
     If r = vbYes Then
        
        PleaseWaitMsg = "Resetting all events ..."
        DoCmd.RunMacro "ShowPleaseWait"

    
        Dim Criteria As String, db As Database, Rs As Recordset
        
        Set db = CurrentDb()
        Set Rs = db.OpenRecordset("SELECT * FROM Heats ORDER BY [E_CODE], [F_LEV] DESC ", dbOpenDynaset)   ' Create dynaset.
        
        Rs.MoveFirst
        PrevE_Code = Rs![E_Code]
        PrevF_Lev = Rs![F_Lev]
        Rs.Edit
        Rs![Completed] = No
        Rs![Status] = evStatus.Current    ' 0=future; 1=active
        Rs.Update
    
        Rs.MoveNext
    
        CurrentlyActive = True
    
        Do Until Rs.EOF
            
            Rs.Edit
            Rs![Completed] = No
            
    
            If Rs![E_Code] = PrevE_Code Then
                
                If PrevF_Lev <> Rs![F_Lev] Then
                    CurrentlyActive = False
                End If
                
                If CurrentlyActive Then
                    Rs![Status] = evStatus.Current   ' 0=future; 1=active
                Else
                    Rs![Status] = evStatus.Future
                End If
    
                PrevE_Code = Rs![E_Code]
                PrevF_Lev = Rs![F_Lev]
    
                Rs.Update
                Rs.MoveNext
    
            Else
                CurrentlyActive = True
                Rs![Status] = evStatus.Current
                PrevE_Code = Rs![E_Code]
                PrevF_Lev = Rs![F_Lev]
                Rs.Update
                Rs.MoveNext
            End If
    
        Loop
    
        Rs.Close
    
        'q = "DELETE DISTINCTROW CompEvents.* FROM CompEvents"
        
        'DoCmd.SetWarnings False
        'DoCmd.RunSQL q
        'DoCmd.SetWarnings False
        DoCmd.RunMacro "ClosePleaseWait"
     End If


Reset_Click_Exit:
    Set db = Nothing
    Exit Sub

Reset_Click_Err:
    MsgBox (Error$)
    GoTo Reset_Click_Exit

End Sub

Private Sub ResetPoints_Click()
    On Error GoTo ResetPoints_Click_Err
    Dim Qry As String, r As Boolean
        
    ' Empty linked table House Points-Extra
    
    r = MsgBox("Resetting all entered extra Team points.  Are you sure you wish to reset all events?", vbYesNo + vbQuestion, "Reset all events")
    If r = vbYes Then
        Qry = "DELETE FROM [House Points-Extra];"
        'DoCmd.SetWarnings False
        'DoCmd.RunSQL qry
        'DoCmd.SetWarnings True
        CurrentDb.Execute (Qry)
    End If
     
     
ResetPoints_Click_Exit:
    Set db = Nothing
    Exit Sub

ResetPoints_Click_Err:
    MsgBox (Error$)
    GoTo ResetPoints_Click_Exit

End Sub
