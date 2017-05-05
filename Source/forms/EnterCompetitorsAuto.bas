Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =8103
    ItemSuffix =115
    Left =3255
    Top =2325
    Right =13620
    Bottom =11880
    HelpContextId =400
    RecSrcDt = Begin
        0x15c96b5db6f2e140
    End
    Caption ="Enter Competitors Automatically"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
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
            Width =850
            Height =850
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
            LabelX =-236
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            BorderLineStyle =0
            Width =187
            Height =187
            LabelX =-236
        End
        Begin OptionGroup
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            OldBorderStyle =0
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
            BackColor =12632256
        End
        Begin ListBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
            BackColor =12632256
        End
        Begin ComboBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
            BackColor =12632256
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
            Height =3396
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =93
                    Width =5613
                    Height =3396
                    BackColor =12440319
                    Name ="Box51"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6236
                    Top =2550
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    Name ="Close"
                    Caption ="Close"
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
                    Left =6236
                    Top =1870
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    HelpContextId =400
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    OverlapFlags =223
                    ColumnCount =2
                    ListWidth =1840
                    Left =2928
                    Top =453
                    Width =1840
                    TabIndex =2
                    BackColor =16777215
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"60\""
                    Name ="Event"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT EventType.ET_Code, EventType.ET_Des FROM EventType ORDER BY EventType.ET_"
                        "Des;"
                    ColumnWidths ="0;1590"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =1927
                            Top =453
                            Width =945
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text94"
                            Caption ="Event:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =223
                    ColumnCount =2
                    ListWidth =1840
                    Left =2926
                    Top =1246
                    Width =1840
                    TabIndex =3
                    BackColor =16777215
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"0\";\"0\""
                    Name ="Selected Age"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Events.Age FROM Events ORDER BY Events.Age;"
                    ColumnWidths ="0;1591"
                    FontName ="Tahoma"
                    EventProcPrefix ="Selected_Age"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =2375
                            Top =1246
                            Width =495
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text96"
                            Caption ="Age:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin OptionButton
                    OverlapFlags =223
                    Left =2108
                    Top =1246
                    TabIndex =4
                    Name ="AllAges"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="Yes"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =957
                            Top =1246
                            Width =915
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text98"
                            Caption ="All Ages:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =2
                    OverlapFlags =215
                    Left =632
                    Top =1983
                    Width =4430
                    Height =970
                    TabIndex =5
                    Name ="CreateHeats"
                    DefaultValue ="2"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =0
                            Left =737
                            Top =1863
                            Width =1084
                            Height =255
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text100"
                            Caption ="Heat creation"
                            FontName ="Tahoma"
                        End
                        Begin OptionButton
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =900
                            Top =2238
                            OptionValue =1
                            Name ="Button102"

                            Begin
                                Begin Label
                                    Visible = NotDefault
                                    BackStyle =0
                                    OverlapFlags =215
                                    TextAlign =1
                                    Left =1173
                                    Top =2210
                                    Width =3799
                                    Height =240
                                    FontWeight =400
                                    Name ="Text103"
                                    Caption ="Fill All Available Heats only"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =900
                            Top =2556
                            OptionValue =2
                            Name ="Button104"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =215
                                    TextAlign =1
                                    Left =1173
                                    Top =2528
                                    Width =3799
                                    Height =240
                                    FontWeight =400
                                    Name ="Text105"
                                    Caption ="Create Heats until all competitors have been added"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =247
                    Left =623
                    Top =1019
                    Width =4472
                    Height =737
                    Name ="Box112"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6236
                    Top =680
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    Name ="AddCompet"
                    Caption ="Add Competitors"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Rectangle
                    BackStyle =0
                    OverlapFlags =247
                    Left =623
                    Top =278
                    Width =4472
                    Height =572
                    Name ="Box114"
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
Option Explicit





Private Sub AddCompet_Click()
                        
    ' Order all competitors in query
    ' Set up array for each house/school, last competitor
    ' Get number of competitors (NumComp)from each house/school
    '
    ' For each event (age/sex division) do
    '   if Create Heats then
    '       While there are more Competitors
    '           Find next heat
    '           If no heat then
    '               Create heat
    '           endif
    '           i=1
    '           while NumComp <> i and more competitors
    '               Get first house
    '               Get next competitor in house/age/sex
    '               While more houses and more competitors and NumComp <> i
    '                   Add to heat
    '                   i = i + 1
    '                   Get next house
    '                   Get next competitor in house/age/sex
    '               wend
    '           wend
    '       wend
    '   else
    '       Find first heat
    '       While there are more heats
    '           i=1
    '           while NumComp <> i and more competitors
    '               Get first house
    '               Get next competitor in house/age/sex
    '               While more houses and more competitors and NumComp <> i
    '                   Add to heat
    '                   i = i + 1
    '                   Get next house
    '                   Get next competitor in house/age/sex
    '               wend
    '           wend
    '           Get next heat
    '       wend
    '   endif
    '

    'Stop

    Dim Crs As Recordset, db As Database, Q As Variant, Ers As Recordset, Hrs As Recordset, HeatRS As Recordset
    Dim TotalHouses As Variant, NumOfLanes As Variant, FLev As Variant, CriteriaHeat As Variant
    Dim CreateHeat As Variant, PointScale As Variant, ProType As Variant, UseTimes As Variant
    Dim Criteria As Variant, Continue As Variant, i As Variant, HouseIndex As Variant, CompLane As Variant
    Dim MoreCompetitors As Variant, ReturnValue As Variant, x As Variant, TotalCompetitors As Variant
    Dim CErs As Recordset
    Dim HC() As HouseComp

    Set db = DBEngine.Workspaces(0).Databases(0)

    Q = "SELECT DISTINCTROW Competitors.PIN, [Surname] & [Gname] & Str([Pin]) AS FullName, Competitors.Age, Competitors.sex, Competitors.H_Code FROM Competitors "
    Q = Q & " ORDER BY [Surname] & [Gname] & Str([Pin])"
    Set Crs = db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.

    Q = "SELECT DISTINCTROW Events.E_Code, Events.Sex, Events.Age, EventType.ET_Des, EventType.EntrantNum, EventType.Lane_Cnt "
    Q = Q & "FROM EventType INNER JOIN Events ON EventType.ET_Code = Events.ET_Code "
    Q = Q & "WHERE EventType.ET_Code=" & Me![Event]
    If Not Me![AllAges] Then
        Q = Q & " AND Events.Age=""" & Me![Selected Age] & """"
    End If
    Set Ers = db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.

    Q = "SELECT DISTINCTROW House.H_Code, House.Include, House.H_ID FROM House "
    Q = Q & "WHERE House.Include=Yes"
    Set Hrs = db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.

    Set CErs = db.OpenRecordset("CompEvents", dbOpenDynaset)   ' Create dynaset.

    Q = "SELECT DISTINCTROW Heats.E_Code, Heats.Heat, Heats.PtScale, Heats.Pro_Type, Heats.UseTimes, Heats.F_Lev, Heats.Completed, Heats.Status FROM Heats"
    'q = q & " ORDER [E_Code],[F_Lev],[Heat]"
    Set HeatRS = db.OpenRecordset(Q, dbOpenDynaset)   ' Create dynaset.

    Hrs.MoveFirst
    TotalHouses = 0
    While Not Hrs.EOF
        TotalHouses = TotalHouses + 1
        ReDim Preserve HC(TotalHouses) As HouseComp
        HC(TotalHouses).H = Hrs!H_Code
        HC(TotalHouses).Hid = Hrs!H_ID
        HC(TotalHouses).c = ""
        Hrs.MoveNext
    Wend
    
    Ers.MoveFirst
    While Not Ers.EOF
        'Stop
        NumOfLanes = Ers![Lane_Cnt]

        ' Determine the Details of each heat to be added
        FLev = DMax("[F_Lev]", "Heats", "[E_Code]=" & Ers!E_Code)
        CriteriaHeat = "[E_Code]=" & Ers!E_Code & " AND [F_Lev]=" & FLev
        HeatRS.FindFirst CriteriaHeat

        If HeatRS.NoMatch Then
            MsgBox ("There must be at least one heat setup for event you want to automatically add competitors into.")
        Else
            CreateHeat = False
            PointScale = HeatRS!PtScale
            ProType = HeatRS!Pro_Type
            UseTimes = HeatRS!UseTimes

            If Me![CreateHeats].Value = 2 Then ' Create Heats option
                'Stop
                Criteria = "val([Age]) " & AgeFilter(Ers!Age) & " AND [Sex]=""" & Ers![Sex] & """"
                TotalCompetitors = DCount("[PIN]", "Competitors", Criteria)
                Crs.FindFirst Criteria
                Continue = True
                ReturnValue = SysCmd(acSysCmdInitMeter, "Adding Competitors to heat ...", TotalCompetitors)
                x = 0
                While Not Crs.NoMatch And Continue   ' This is the beggining of a new heat
                                                     ' If there are still competitors then we want to add a new heat
                                                     ' It will stop when a whole heat is left empty
                    
                    ReturnValue = SysCmd(acSysCmdUpdateMeter, x)
                    Continue = False
                    i = 1 'DCount("[F_Lev]", "CompEvents", "[E_Code]=" & ERS!E_Code & " AND [F_Lev]=" & Flev & " AND [Heat]=" & HeatRS!Heat) + 1
                    MoreCompetitors = True
                    While (i <= NumOfLanes Or NumOfLanes = 0) And MoreCompetitors 'And Not CRS.nomatch
                        'Stop
                        MoreCompetitors = False
                        HouseIndex = 1
                        While (HouseIndex <= TotalHouses) And (i <= NumOfLanes Or NumOfLanes = 0)
                            Criteria = "val([Age]) " & AgeFilter(Ers!Age) & " AND [Sex]=""" & Ers![Sex] & """"
                            Criteria = Criteria & " AND [FullName]>""" & HC(HouseIndex).c & """ AND [H_Code]=""" & HC(HouseIndex).H & """"
                            Crs.FindFirst Criteria
                            If Not Crs.NoMatch Then
                                MoreCompetitors = True
                                x = x + 1
                                If CreateHeat Then
                                    'Stop
                                     HeatRS.AddNew
                                     HeatRS!E_Code = Ers!E_Code
                                     HeatRS!Heat = DMax("[Heat]", "Heats", "[E_Code]=" & Ers!E_Code & " AND [F_Lev]=" & FLev) + 1
                                     HeatRS!PtScale = PointScale
                                     HeatRS!Pro_Type = ProType
                                     HeatRS!UseTimes = UseTimes
                                     HeatRS!F_Lev = FLev
                                     HeatRS!Status = 1
                                     HeatRS!Completed = False
                                     HeatRS.Update
                                     HeatRS.MoveLast
                                     CreateHeat = False
                                End If
                                
                                i = i + 1
                                ' Add Competitor to heat
                                'Stop
                                CompLane = Calculate_Competitor_Lane(Ers!E_Code, FLev, HC(HouseIndex).Hid, HeatRS!Heat)
                                CErs.AddNew

                                CErs!PIN = Crs!PIN
                                CErs!E_Code = Ers!E_Code
                                CErs!F_Lev = FLev
                                CErs!Heat = HeatRS!Heat
                                CErs!Lane = CompLane
                                CErs!Place = 0
                                CErs!Result = "0"
                                CErs!nResult = 0
                                CErs!Points = 0
                                CErs.Update
                                  
                                'q = "INSERT INTO CompEvents (PIN, E_Code, F_Lev, Heat, Lane, Place, Result, nResult, Points)"
                                'q = q & " VALUES (" & CRS!PIN & "," & ERS!E_Code & "," & Flev & "," & HeatRS!Heat & "," & CompLane & ", 0, ""0"",0, 0 )"
                                'DoCmd SetWarnings False
                                'DoCmd RunSQL q
                                'DoCmd SetWarnings True

                                HC(HouseIndex).c = Crs!Fullname
                                Continue = True
                            End If
                            HouseIndex = HouseIndex + 1
                        Wend
                    Wend
                    'Stop
                    CriteriaHeat = "[E_Code]=" & Ers!E_Code & " AND [F_Lev]=" & FLev & " AND [Heat]=" & (HeatRS!Heat + 1)
                    HeatRS.FindFirst CriteriaHeat
                    If HeatRS.NoMatch Then
                        CreateHeat = True
                    Else
                        CreateHeat = False
                    End If
                Wend
    
            End If
        End If
        Ers.MoveNext
        For HouseIndex = 1 To TotalHouses
            HC(HouseIndex).c = ""
        Next
    Wend

    Crs.Close
    Ers.Close
    Hrs.Close
    HeatRS.Close
    ReturnValue = SysCmd(acSysCmdRemoveMeter)
    

    ' For each event (age/sex division) do
    '   if Create Heats then
    '       While there are more Competitors
    '           Find next heat
    '           If no heat then
    '               Create heat
    '           endif
    '           i=1
    '           while NumComp <> i and more competitors
    '               Get first house
    '               Get next competitor in house/age/sex
    '               While more houses and more competitors and NumComp <> i
    '                   Add to heat
    '                   i = i + 1
    '                   Get next house
    '                   Get next competitor in house/age/sex
    '               wend
    '           wend
    '       wend
    '   else
    '       Find first heat
    '       While there are more heats
    '           i=1
    '           while NumComp <> i and more competitors
    '               Get first house
    '               Get next competitor in house/age/sex
    '               While more houses and more competitors and NumComp <> i
    '                   Add to heat
    '                   i = i + 1
    '                   Get next house
    '                   Get next competitor in house/age/sex
    '               wend
    '           wend
    '           Get next heat
    '       wend
    '   endif


    '


End Sub

Private Sub AllAges_AfterUpdate()

    If Me![AllAges].Value = False Then
        Me![Selected Age].enabled = True
    Else
        Me![Selected Age].enabled = False
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
