Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    GridY =10
    Width =9127
    ItemSuffix =46
    Left =-17970
    Top =2655
    Right =-8205
    Bottom =10665
    HelpContextId =140
    RecSrcDt = Begin
        0x842055290fcde140
    End
    RecordSource ="MiscellaneousLocal"
    Caption ="Create Carnival Disks"
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
        Begin Line
            BorderLineStyle =0
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
        Begin Section
            CanGrow = NotDefault
            Height =7088
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =565
                    Top =1765
                    Width =4080
                    Height =5039
                    BorderColor =12632256
                    Name ="Embedded0"
                    SourceObject ="Form.ExportTextSF1"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =573
                            Top =1195
                            Width =4185
                            Height =240
                            FontWeight =700
                            Name ="Text1"
                            Caption ="Team to generate Carnival Disk(s) for:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =223
                    Left =113
                    Top =1020
                    Width =5032
                    Height =5959
                    Name ="Box2"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =2010
                    Top =1530
                    Width =1845
                    Height =225
                    Name ="Text4"
                    Caption ="Team"
                    FontName ="Tahoma"
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =5329
                    Top =1027
                    Width =2046
                    Height =843
                    TabIndex =1
                    Name ="SexFormatOB"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    ControlTipText ="How do you want the gender specified?"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5449
                            Top =907
                            Width =922
                            Height =255
                            BackColor =-2147483633
                            Name ="Text7"
                            Caption ="Sex Format"
                            FontName ="Tahoma"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =5561
                            Top =1219
                            OptionValue =1
                            BorderColor =12632256
                            Name ="Field9"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    Left =5907
                                    Top =1191
                                    Width =930
                                    Height =240
                                    Name ="Text10"
                                    Caption ="Boys / Girls"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =5561
                            Top =1537
                            OptionValue =2
                            BorderColor =12632256
                            Name ="Field11"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    Left =5907
                                    Top =1509
                                    Width =1140
                                    Height =240
                                    Name ="Text12"
                                    Caption ="Male / Female"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =5329
                    Top =2161
                    Width =2039
                    Height =828
                    TabIndex =2
                    Name ="HeatFormatOB"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    ControlTipText ="How do you want the heats specified?"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5449
                            Top =2041
                            Width =1001
                            Height =255
                            BackColor =-2147483633
                            Name ="Text14"
                            Caption ="Heat Format"
                            FontName ="Tahoma"
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =5561
                            Top =2359
                            OptionValue =1
                            BorderColor =12632256
                            Name ="Field16"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    Left =5911
                                    Top =2333
                                    Width =975
                                    Height =240
                                    Name ="Text17"
                                    Caption ="1, 2, 3, 4 etc"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =5561
                            Top =2677
                            OptionValue =2
                            BorderColor =12632256
                            Name ="Field18"

                            Begin
                                Begin Label
                                    BackStyle =0
                                    OverlapFlags =87
                                    Left =5911
                                    Top =2648
                                    Width =1050
                                    Height =240
                                    Name ="Text19"
                                    Caption ="A, B, C, D etc"
                                    FontName ="Tahoma"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7785
                    Top =285
                    Width =1134
                    Height =630
                    FontSize =8
                    TabIndex =3
                    ForeColor =32768
                    Name ="CreateBut"
                    Caption ="Create Disks"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Create the carnival disks."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =7785
                    Top =6390
                    Width =1134
                    Height =510
                    FontSize =8
                    TabIndex =4
                    Name ="CloseBut"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =283
                    Top =447
                    Width =6336
                    Height =255
                    TabIndex =5
                    BorderColor =12632256
                    Name ="Path"
                    ControlSource ="ExportLocation"
                    DefaultValue ="\"A:\\\""
                    FontName ="Tahoma"
                    ControlTipText ="Enter the path where you want to create the carnival (don't include the file nam"
                        "e)."

                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =170
                    Top =277
                    Width =7189
                    Height =585
                    Name ="Box24"
                End
                Begin Label
                    OverlapFlags =247
                    Left =223
                    Top =113
                    Width =2325
                    Height =255
                    BackColor =-2147483633
                    Name ="Text25"
                    Caption ="Full Path for Carnival File"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5555
                    Top =3558
                    TabIndex =6
                    BorderColor =12632256
                    Name ="PlainText"
                    DefaultValue ="Yes"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            Left =5818
                            Top =3509
                            Width =1440
                            Height =225
                            Name ="Text33"
                            Caption ="Plain Text"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =3
                    BackStyle =0
                    OverlapFlags =255
                    Left =5329
                    Top =3339
                    Width =2029
                    Height =1260
                    Name ="Box34"
                End
                Begin Label
                    OverlapFlags =247
                    Left =5382
                    Top =3175
                    Width =975
                    Height =255
                    BackColor =-2147483633
                    Name ="Text35"
                    Caption ="Disk Format:"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5555
                    Top =3854
                    TabIndex =7
                    BorderColor =12632256
                    Name ="Excel"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =247
                            Left =5818
                            Top =3861
                            Width =1440
                            Height =225
                            Name ="Text38"
                            Caption ="Plain Text (CSV)"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =7785
                    Top =5685
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    HelpContextId =140
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =6705
                    Top =435
                    Width =576
                    Height =351
                    TabIndex =9
                    Name ="OpenFolderBut"
                    Caption ="Command41"
                    StatusBarText ="Push the button to select the path to place the carnival files."
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
                        0x000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Push the button to select the path to place the carnival files."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =3
                    Left =7540
                    Width =0
                    Height =7088
                    Name ="Line43"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =5555
                    Top =4214
                    TabIndex =10
                    BorderColor =12632256
                    Name ="RTF"
                    DefaultValue ="No"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =0
                            OverlapFlags =247
                            Left =5818
                            Top =4221
                            Width =1560
                            Height =225
                            Name ="Label45"
                            Caption ="Rich Text Format "
                            FontName ="Tahoma"
                        End
                    End
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

Private Sub CreateBut_Click()
On Error GoTo Err_CreateBut_Click

    
    If Me.Excel = False And Me.PlainText = False And Me.RTF = False Then
        MsgBox "You must choose a format for the Carnival file.", vbInformation
        GoTo Exit_CreateBut_Click
    End If

    Dim Criteria As String, Db As Database, rs As Recordset
    Dim ETrs As Recordset, TTRS As Recordset

    Set Db = CurrentDb()
    Set rs = Db.OpenRecordset("House", dbOpenDynaset)   ' Create dynaset.

    GoSub Fill_Num_Of_Entrants

    Criteria = "[Include]=Yes and [Flag]= Yes"
    rs.FindFirst Criteria    ' Find first occurrence.
    
    Do Until rs.NoMatch  ' Loop until no matching records.
        
        HouseName = rs!H_NAme

        Mesg = "Please insert disk for " & HouseName & ".  Do you wish to create this disk?"
        Response = MsgBox(Mesg, 35, "Next Disk")

        If Response = 6 Then 'Yes
            ChosenFinalHouse = rs!H_Code
            GoSub Create_Carnival_Disk
        ElseIf Response = 2 Then
            GoTo Abort_DiskCreate
        End If
        
        rs.FindNext Criteria ' Find next occurrence.

    Loop                            ' End of loop.

    MsgBox "The Carnival disk creation process complete.", vbInformation


Abort_DiskCreate:
    rs.Close

Exit_CreateBut_Click:
    Set Db = Nothing
    Exit Sub

Err_CreateBut_Click:
    MsgBox Error$, vbCritical
    Resume Exit_CreateBut_Click
    

Create_Carnival_Disk:

    If Me.PlainText = True Then
        FullFileName = [Path] & ChosenFinalHouse & ".TXT"
        DoCmd.TransferText acExportDelim, "Create Carnival Disk", "Carnival Disk Export", FullFileName, False
    End If

    If Me.Excel = True Then
        FullFileName = [Path] & ChosenFinalHouse & ".CSV"
        DoCmd.TransferText acExportDelim, "Create Carnival Disk", "Carnival Disk Export", FullFileName, True
    End If

    If Me.RTF = True Then
        FullFileName = [Path] & ChosenFinalHouse & ".RTF"
        DoCmd.OutputTo acOutputQuery, "Carnival Disk Export", acFormatRTF, FullFileName, False
    End If

  Return
    

Fill_Num_Of_Entrants:
    
    Set ETrs = Db.OpenRecordset("EventType", dbOpenDynaset)   ' Create dynaset.
    Set TTRS = Db.OpenRecordset("Temporary Table", dbOpenDynaset)   ' Create dynaset.

    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE DISTINCTROW [Temporary Table].Field1 FROM [Temporary Table]"
    DoCmd.SetWarnings True
    
    ETrs.MoveFirst

    While Not ETrs.EOF

        NumOfEntrants = ETrs!EntrantNum

        For i = 1 To NumOfEntrants
            TTRS.AddNew
            TTRS![ET_Code] = ETrs![ET_Code]
            TTRS!Field1 = i
            TTRS.Update

        Next i
        
        ETrs.MoveNext
    Wend

    ETrs.Close
    TTRS.Close

    Return

End Sub

Private Sub Form_Load()

    SexFormatOB_AfterUpdate
    HeatFormatOB_AfterUpdate

End Sub

Private Sub HeatFormatOB_AfterUpdate()

    If [HeatFormatOB] = "2" Then
        HeatFormat = "ABCD"
    Else
        HeatFormat = "1234"
    End If

End Sub

Private Sub SexFormatOB_AfterUpdate()

    If [SexFormatOB] = "1" Then
        SexFormat = "Boys/Girls"
    Else
        SexFormat = "Male/Female"
    End If

End Sub

Private Sub OpenFolderBut_Click()
On Error GoTo Err_OpenFolderBut_Click


    Dim n As Variant

    'Me!ctlCommonDialog.DialogTitle = "Choose folder for HTML files"
    'Me!ctlCommonDialog.FileName = "LocateFolder"
    'Me!ctlCommonDialog.ShowSave
    
    
    n = BrowseFolder("Locate web folder") 'Me!ctlCommonDialog.FileName
    If n <> "" Then
        Me![Path] = Trim(n)
    End If


Exit_OpenFolderBut_Click:
    Exit Sub

Err_OpenFolderBut_Click:
    MsgBox Err.Description
    Resume Exit_OpenFolderBut_Click
    
End Sub
