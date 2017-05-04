Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DataEntry = NotDefault
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridX =5
    GridY =5
    Width =5329
    ItemSuffix =5
    Left =3255
    Top =2325
    Right =12645
    Bottom =5265
    RecSrcDt = Begin
        0x6de96f5db6f2e140
    End
    Caption ="Copy Event"
    HelpFile ="SportsAdmin.chm"
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
        Begin Section
            Height =1474
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OverlapFlags =93
                    Left =2931
                    Top =229
                    Width =2046
                    Name ="Description"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =3
                            Left =226
                            Top =226
                            Width =2595
                            Height =240
                            Name ="Text1"
                            Caption ="Enter the Name of the new Event:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4081
                    Top =907
                    Width =885
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="CreateBut"
                    Caption ="Create"
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
                    Left =453
                    Top =907
                    Width =915
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="CancelBut"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    SpecialEffect =1
                    BackStyle =0
                    OverlapFlags =247
                    Left =113
                    Top =113
                    Width =5102
                    Height =567
                    Name ="Box4"
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

Private Sub CancelBut_Click()
On Error GoTo Err_CancelBut_Click


    DoCmd.Close

Exit_CancelBut_Click:
    Exit Sub

Err_CancelBut_Click:
    MsgBox Error$
    Resume Exit_CancelBut_Click
    
End Sub

Private Sub CreateBut_Click()

    Dim db As Database, Rs As Recordset, Ers As Recordset, Hrs As Recordset
    Set db = DBEngine.Workspaces(0).Databases(0)
    Dim Rept As Variant, Q As Variant, OldET_Code As Variant, NewET_Code As Variant, Criteria As Variant
    Dim E_Code As Variant, Sex As Variant, Age As Variant, Record As Variant, Include As Variant
    Dim BMark As Variant, NewE_Code As Variant
    
On Error GoTo Err_CreateBut_Click

    ' Get New Event type Description and check

  If Not IsNull([Description]) Then
    If IsNull(DLookup("[ET_Des]", "EventType", "[ET_Des] = """ & [Description] & """")) Then
        
        If Me.OpenArgs = "ADD" Then
            Rept = DLookup("[R_Code]", "ReportTypes")
            Q = "INSERT INTO EventType ( ET_Des, EntrantNum, R_Code, Include, Lane_Cnt, Units) "
            Q = Q & "Values (""" & [Description] & """,1, " & Rept & ", YES,0,'Secs')"

            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
            
            Set Rs = db.OpenRecordset("EventType", dbOpenDynaset)   ' Create Recordset.
        
            Rs.MoveLast
            NewET_Code = Rs!ET_Code
        
            Rs.Close
            
            GlobalVariable = NewET_Code
            GlobalCancel = False

            DoCmd.Close
            GoTo Exit_CreateBut_Click

        End If
        
        OldET_Code = Forms![EventTypeSummary]![Summary]

        ' Add to EventType Table
        
        Q = "INSERT INTO EventType ( ET_Des, Units, Lane_Cnt, R_Code, Include, EntrantNum ) "
        Q = Q & "SELECT DISTINCTROW """ & [Description] & """ AS Expr2, EventType.Units, EventType.Lane_Cnt, EventType.R_Code, EventType.Include, EventType.EntrantNum "
        Q = Q & "FROM EventType WHERE EventType.ET_Code= " & OldET_Code

        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True


        Set Rs = db.OpenRecordset("EventType", dbOpenDynaset)   ' Create Recordset.
    
        Rs.MoveLast
        NewET_Code = Rs!ET_Code
    
        Rs.Close
        
        ' Add to Final_Lev Table
        Q = "INSERT INTO Final_Lev ( ET_Code, F_Lev, NoHeats, PtScale, ProType, UseTimes ) "
        Q = Q & "SELECT DISTINCTROW " & NewET_Code & " AS Expr1, Final_Lev.F_Lev, Final_Lev.NoHeats, Final_Lev.PtScale, Final_Lev.ProType, Final_Lev.UseTimes "
        Q = Q & "FROM Final_Lev WHERE Final_Lev.ET_Code= " & OldET_Code

        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True

        ' Add to Lane Promotion Allocation Table
        Q = "INSERT INTO [Lane Promotion Allocation] ( ET_Code, Place, Lane ) "
        Q = Q & "SELECT DISTINCTROW " & NewET_Code & " AS Expr1, [Lane Promotion Allocation].Place, [Lane Promotion Allocation].Lane "
        Q = Q & "FROM [Lane Promotion Allocation] WHERE [Lane Promotion Allocation].ET_Code= " & OldET_Code

        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True

        Q = "INSERT INTO [Lane Template] ( ET_Code, Lanes ) "
        Q = Q & "SELECT DISTINCTROW " & NewET_Code & " AS Expr1, [Lane Template].Lanes "
        Q = Q & "FROM [Lane Template] WHERE [Lane Template].ET_Code=" & OldET_Code

        DoCmd.SetWarnings False
        DoCmd.RunSQL Q
        DoCmd.SetWarnings True

        ' Add to Events Table
        
        Set Ers = db.OpenRecordset("Events", dbOpenDynaset)   ' Create Recordset.
        Set Hrs = db.OpenRecordset("Heats", dbOpenDynaset)   ' Create Recordset.
        
        Criteria = "[ET_Code] = " & OldET_Code
        Ers.FindFirst Criteria

        Do Until Ers.NoMatch  ' Loop until no matching records.
            
            E_Code = Ers!E_Code
            Sex = Ers!Sex
            Age = Ers!Age
            Record = Ers!Record
            Include = Ers!Include
            
            Ers.AddNew
            
            Ers!ET_Code = NewET_Code
            Ers!Sex = Sex
            Ers!Age = Age
            Ers!Record = " "
            Ers!Include = Include
            
            Ers.Update        ' Save changes.

            BMark = Ers.Bookmark         ' Obtain bookmark.
    
            Ers.MoveLast
            NewE_Code = Ers!E_Code

            Ers.Bookmark = BMark

            ' Add Heats
            Q = "INSERT INTO Heats ( E_Code, Heat, PtScale, E_Number, E_Time, F_Lev, Pro_Type, UseTimes, Completed, Status ) "
            Q = Q & "SELECT DISTINCTROW " & NewE_Code & " AS Expr1, Heats.Heat, Heats.PtScale, NULL, Heats.E_Time, Heats.F_Lev, Heats.Pro_Type, Heats.UseTimes, Heats.Completed, Heats.Status "
            Q = Q & "FROM Heats WHERE Heats.E_Code= " & E_Code

            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
            
            Ers.FindNext Criteria ' Locate next record.

        Loop
        
        
        Ers.Close
        GlobalVariable = NewET_Code
        GlobalCancel = False

        DoCmd.Close
          
    Else
        MsgBox ("The event that you typed is already present.  Please choose a different event name.")
    End If

  End If

Exit_CreateBut_Click:
    
    DoCmd.SetWarnings True
    Exit Sub

Err_CreateBut_Click:
    MsgBox Error$
    Resume Exit_CreateBut_Click
    
End Sub

Private Sub Form_Load()

    If Me.OpenArgs = "ADD" Then
        Me.[Caption] = "Add Event"
    Else
        Me.[Caption] = "Copy Event"
    End If
        
End Sub
