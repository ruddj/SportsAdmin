Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DataEntry = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =5867
    ItemSuffix =13
    Left =2790
    Top =2295
    Right =9255
    Bottom =6495
    HelpContextId =30
    RecSrcDt = Begin
        0x8f98da87edc6e140
    End
    Caption ="Copy Event"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
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
        Begin TextBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin Section
            Height =2834
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    Left =340
                    Top =737
                    Width =4701
                    Name ="Description"
                    FontName ="Tahoma"
                    ControlTipText ="Enter a descriptive name for the carnival."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =345
                            Top =450
                            Width =4740
                            Height =240
                            Name ="Text1"
                            Caption ="Enter a descriptive name for the new Carnival:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4847
                    Top =2182
                    Width =885
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="Button"
                    Caption ="Create"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Create the carnival."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =226
                    Top =2211
                    Width =915
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =3
                    Name ="CancelBut"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Cancel this operation."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =3
                    OverlapFlags =93
                    BackStyle =0
                    Left =340
                    Top =1360
                    Width =4701
                    TabIndex =4
                    BackColor =12632256
                    Name ="FullFileName"
                    FontName ="Tahoma"
                    ControlTipText ="Click the 'Open' button to enter or select the name of a carnival."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =345
                            Top =1080
                            Width =4785
                            Height =240
                            Name ="FileNameText"
                            Caption ="Click the 'Open' button to select a file name for the new carnival."
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =255
                    Left =219
                    Top =350
                    Width =5537
                    Height =1737
                    Name ="Box7"
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontFamily =34
                    Left =5100
                    Top =1335
                    Width =576
                    Height =351
                    FontSize =7
                    FontWeight =400
                    TabIndex =1
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
                        0x0000000000000000
                    End
                    FontName ="Tahoma"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Click to locate or enter a file name for this carnival."

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
Option Explicit

Private Function AddCarnival() As Variant

    On Error GoTo Err_AddCarnival
    Dim MyDb As Database, ITable As Recordset, SpecifiedPath As Variant, TT As TableDef, FTable As Recordset
    Dim DataExists As Variant, MyWS As Workspace, CPath  As Variant, AskUser  As Variant
    Dim Result As Variant, ReturnVal As Variant, Db As Database
    Dim NewDir As String, OldDB As String, NextCarn As String
    AskUser = False
    
    If IsNull(Me![FullFileName]) Then
        MsgBox ("You must enter a valid carnival file.")
    Else
        Result = Me![FullFileName]
        Result = Trim$(Result)
        If InStr(ReverseString(CStr(Result)), "\") <> 0 Then
            Set MyWS = DBEngine.Workspaces(0)
            Set MyDb = CurrentDb()
            Set Db = MyWS.OpenDatabase(Result)
            Set ITable = MyDb.OpenRecordset("SELECT * FROM [Inventory Attached Tables] Where [IF ID] = 2;")
            Do Until ITable.EOF
                If ITable![Table Name] = "CompetitorsOrdered" Then
                    ' Do nothing
                Else
                    Set TT = MyDb.TableDefs(ITable![Table Name])
                End If
                ITable.MoveNext
            Loop
            ITable.Close
            
            NextCarn = GetCarnivalFile(Result)
            NewDir = GetCarnivalRelDir(Result)
                
            DoCmd.SetWarnings False
            DoCmd.RunSQL "INSERT INTO Carnivals ( Carnival, Filename, [Relative Directory], Available ) SELECT DISTINCTROW """ & Me.[Description] & """ AS E1, """ & NextCarn & """ AS E2, """ & NewDir & """ AS E3, True AS Expr4;"
            DoCmd.SetWarnings True
            AskUser = True
            
            
        
        End If
    End If
    
Exit_AddCarnival:
    Set MyDb = Nothing
    DoCmd.SetWarnings True
    AddCarnival = AskUser
    Exit Function
    
Err_AddCarnival:
    MsgBox "This file does not meet the specified format for a carnival file.  Problem: " & Error$, vbExclamation, "Message"
    Resume Exit_AddCarnival
    
End Function



Private Sub Button_Click()
    
    On Error GoTo Err_Button_Click
    'Stop
    Dim RelDir As Variant, FullDir As Variant

    'Me.[Caption] = "Add Carnival"

    Dim NewName As String, NewDir As Variant, NextCarn As String, OldDB  As String, Result As Variant
    Result = False
    If IsNull(Me.[Description]) Then
        MsgBox "Enter the name of the carnival to continue.", vbExclamation, "Message"
    Else
        NewName = Trim$(Me.[Description])
        If NewName = "" Then
            Me.[Description] = NewName
            MsgBox "Enter the name of the carnival to continue.", vbExclamation, "Message"
        Else
            If Me.[Caption] = "Rename Carnival" Then
                If Not IsNull(DLookup("[Carnival]", "Carnivals", "[Carnival] = """ & NewName & """")) Then
                    MsgBox "There is already a carnival by this name. Choose another name.", vbExclamation, "Message"
                Else
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "UPDATE DISTINCTROW Carnivals SET Carnivals.Carnival = """ & NewName & """ WHERE (Carnivals.Carnival=""" & Forms![Carnivals Maintain]![List] & """);"
                    DoCmd.SetWarnings True
                    Result = True
                End If
            ElseIf Me.[Caption] = "New Carnival" Then
              If IsNull(Me![FullFileName]) Then
                MsgBox ("You must enter a filename for the new carnival.")
              Else
                
                RelDir = GetCarnivalRelDir(Me![FullFileName])
                FullDir = GetCarnivalFullDir(Me![FullFileName])

                ChDrive Left$(FullDir, 1)
                
                On Error Resume Next
                
                MkDir Left$(FullDir, Len(FullDir) - 1)

                If Err = 75 Then    ' Check if directory exists.
                    ' directory already exists."
                ElseIf Err Then
                    MsgBox ("Carnival directory could not be created.")
                    GoTo Err_Button_Click
                End If
                
                NextCarn = GetCarnivalFile(Me![FullFileName])

                If Make_File(FullDir & NextCarn) Then
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "INSERT INTO Carnivals ( Carnival, Filename, [Relative Directory], Available ) SELECT DISTINCTROW """ _
                      & Me.[Description] & """ AS E1, """ & NextCarn & """ AS E2, """ & RelDir & """ AS E3, True AS Expr4;"
                    DoCmd.SetWarnings True
                    Result = True
                    Response = MsgBox("Your new carnival has been created and will appear in the carnival list.  Double-click it to start working on it.", vbInformation)
                End If
              End If
            ElseIf Me.[Caption] = "Copy Carnival" Then
              If IsNull(Me![FullFileName]) Then
                MsgBox "You must enter a filename for the new carnival.", vbInformation
              Else
                ''NewDir = DLookup("[Relative Directory]", "Carnivals", "[Carnival] =""" & Forms![Carnivals Maintain]![List] & """")
                ''OldDB = DBpath() & NewDir & DLookup("[Filename]", "Carnivals", "[Carnival] =""" & Forms![Carnivals Maintain]![List] & """")
                ''NextCarn = NextCarnival()

                OldDB = CarnivalDir(DLookup("[Relative Directory]", "Carnivals", "[Carnival] =""" & Forms![Carnivals Maintain]![List] & """"))
                OldDB = OldDB & DLookup("[Filename]", "Carnivals", "[Carnival] =""" & Forms![Carnivals Maintain]![List] & """")

                NextCarn = GetCarnivalFile(Me![FullFileName])
                'FullDir = GetCarnivalFullDir(Me![FullFileName])
                RelDir = GetCarnivalRelDir(Me![FullFileName])
                
                Call CloseAlwaysOpenRS
                
                DBEngine.CompactDatabase OldDB, Me![FullFileName], dbLangGeneral
                DoCmd.SetWarnings False
                DoCmd.RunSQL "INSERT INTO Carnivals ( Carnival, Filename, [Relative Directory], Available ) SELECT DISTINCTROW """ & Me.[Description] & """ AS E1, """ & NextCarn & """ AS E2, """ & RelDir & """ AS E3, True AS Expr4;"
                DoCmd.SetWarnings True
                Result = True
                Call OpenAlwaysOpenRS
              End If
            ElseIf Me.[Caption] = "Add Carnival" Then
                If AddCarnival() Then
                    Result = True
                End If
            Else
                MsgBox "Press cancel to exit this form.", vbExclamation, "Message"
            End If
        End If
    End If
    If Result Then
        CancelBut_Click
    End If
Exit_Button_Click:
    DoCmd.SetWarnings True
    Exit Sub
Err_Button_Click:
    MsgBox Error$, vbCritical
    Resume Exit_Button_Click
End Sub

Private Sub CancelBut_Click()
On Error GoTo Err_CancelBut_Click


    DoCmd.Close

Exit_CancelBut_Click:
    Exit Sub

Err_CancelBut_Click:
    MsgBox Error$
    Resume Exit_CancelBut_Click
    
End Sub

Private Sub Form_Load()

    If Me.OpenArgs = "NEW" Then
        Me.[Caption] = "New Carnival"
        Me![Button].Caption = "Create"
'        Me![FullFileName].visible = False
'        Me![Locate].visible = False
'        Me![FileNameText].visible = False

    ElseIf Me.OpenArgs = "RENAME" Then
        Me.[Caption] = "Rename Carnival"
        Me![Button].Caption = "Rename"
        Me![Locate].visible = False
        Me![FullFileName].visible = False
        Me!Description = GlobalVariable
    
    ElseIf Me.OpenArgs = "COPY" Then
        Me.[Caption] = "Copy Carnival"
        Me![Button].Caption = "Copy"
    
    ElseIf Me.OpenArgs = "DELETE" Then
        Me.[Caption] = "Delete Carnival"
        Me![Button].Caption = "Delete"
        Me![Locate].visible = False
        Me![FullFileName].visible = False
    
    ElseIf Me.OpenArgs = "ADD" Then
        Me.[Caption] = "Add Carnival"
        Me![Button].Caption = "Add"
    End If
        
End Sub

Private Sub Locate_Click()
On Error GoTo Err_Locate_Click

    Dim n As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    If Me.[Caption] = "New Carnival" Then
      strFilter = ahtAddFilterItem(strFilter, "Carnival Files (*.accdb)", "*.accdb")
      n = ahtCommonFileOpenSave(InitialDir:="", OpenFile:=False, _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate Carnival File")
    Else
      strFilter = ahtAddFilterItem(strFilter, "Carnival Files (*.accdb, *.mdb)", "*.accdb;*.mdb")
      strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
      n = ahtCommonFileOpenSave(InitialDir:="", _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate Carnival File")
    End If


    If n <> "" Then
        Me![FullFileName] = Trim(n)
    End If

Exit_Locate_Click:
    Exit Sub

Err_Locate_Click:
    MsgBox Error$
    Resume Exit_Locate_Click
    
End Sub
