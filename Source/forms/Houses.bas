Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    GridX =20
    GridY =20
    Width =8280
    ItemSuffix =49
    Left =1575
    Top =2250
    Right =11670
    Bottom =10635
    HelpContextId =40
    Filter ="[H_ID]=2"
    RecSrcDt = Begin
        0x7030d5b911cde140
    End
    RecordSource ="SELECT DISTINCTROW House.H_Code, House.H_NAme, House.Details, House.Lane, House."
        "CompPool, House.H_ID, House.Include FROM House ORDER BY House.H_NAme;"
    Caption ="Team Details"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
            Height =4560
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1514
                    Top =141
                    Width =3150
                    BorderColor =12632256
                    Name ="H_Code"
                    ControlSource ="H_Code"
                    Format =">"
                    StatusBarText ="House / School Code ie. Asher, COC, Beaudesert, Australia, Individual?)"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =566
                            Top =141
                            Width =879
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text15"
                            Caption ="Code:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1519
                    Top =519
                    Width =3120
                    TabIndex =1
                    BorderColor =12632256
                    Name ="H_NAme"
                    ControlSource ="H_NAme"
                    StatusBarText ="House / School Name"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =679
                            Top =519
                            Width =765
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text17"
                            Caption ="Name:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    ScrollBars =2
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1513
                    Top =1937
                    Width =4845
                    Height =900
                    TabIndex =5
                    BorderColor =12632256
                    Name ="Details"
                    ControlSource ="Details"
                    StatusBarText ="Address etc."
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =679
                            Top =1937
                            Width =765
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text21"
                            Caption ="Details:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    ListWidth =2785
                    Left =1515
                    Top =897
                    Width =3130
                    Height =240
                    TabIndex =2
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Field22"
                    ControlSource ="House.HT_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT HouseTypes.HT_Code, HouseTypes.Desc FROM HouseTypes;"
                    ColumnWidths ="0;2161"
                    DefaultValue ="1"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =283
                            Top =897
                            Width =1155
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text23"
                            Caption ="Team Type:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6916
                    Top =188
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =6
                    ForeColor =8404992
                    Name ="Button34"
                    Caption ="Allocate Lanes"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1510
                    Top =1275
                    Width =3135
                    TabIndex =3
                    BorderColor =12632256
                    Name ="Field35"
                    ControlSource ="CompPool"
                    Format =">"
                    StatusBarText ="House / School Code ie. Asher, COC, Beaudesert, Australia, Individual?)"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =121
                            Top =1282
                            Width =1290
                            Height =270
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text36"
                            Caption ="Competitor Pool:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    TextFontFamily =34
                    Left =6912
                    Top =3960
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =8
                    Name ="Button28"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1542
                    Top =1653
                    TabIndex =4
                    BorderColor =12632256
                    Name ="Include"
                    ControlSource ="Include"
                    DefaultValue ="Yes"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =623
                            Top =1653
                            Width =810
                            Height =240
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text39"
                            Caption ="Include"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6916
                    Top =818
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =7
                    ForeColor =8404992
                    Name ="Extra"
                    Caption ="Allocate Extra Points"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Subform
                    OverlapFlags =85
                    Left =144
                    Top =3240
                    Width =6193
                    Height =1233
                    TabIndex =9
                    BorderColor =12632256
                    Name ="Embedded2"
                    SourceObject ="Form.House Points-Extra SF"
                    LinkChildFields ="H_ID"
                    LinkMasterFields ="H_ID"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =144
                    Top =2952
                    Width =1140
                    Height =240
                    BackColor =-2147483633
                    Name ="Text45"
                    Caption ="Extra Points"
                    FontName ="Tahoma"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =85
                    Left =6624
                    Width =0
                    Height =4560
                    Name ="Line46"
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

Private Sub Button34_Click()
On Error GoTo Err_Button34_Click

    Call SaveRecord_Click
    
    Dim DocName As String
    Dim LinkCriteria As String

    DocName = "Lanes"
    DoCmd.OpenForm DocName, , , LinkCriteria, , acDialog

Exit_Button34_Click:
    Exit Sub

Err_Button34_Click:
    MsgBox Error$
    Resume Exit_Button34_Click
    
End Sub

Private Sub Button37_Click()
On Error GoTo Err_Button37_Click


    DoCmd.GoToRecord , , A_NEWREC

Exit_Button37_Click:
    Exit Sub

Err_Button37_Click:
    MsgBox Error$
    Resume Exit_Button37_Click
    
End Sub

Private Sub Extra_Click()

    Call SaveRecord_Click
    DoCmd.OpenForm "House Points-Extra", , , , , acDialog

End Sub

Private Sub Form_Load()

    If OpenFormType = "ADD" Then
        'Me.DefaultEditing = 1
        DoCmd.GoToRecord , , A_NEWREC
    Else
        'Me.DefaultEditing = 4
    End If

    Me!H_Code.SetFocus
        
End Sub

Private Sub Command47_Click()
On Error GoTo Err_Command47_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_Command47_Click:
    Exit Sub

Err_Command47_Click:
    MsgBox Err.Description
    Resume Exit_Command47_Click
    
End Sub
Private Sub SaveRecord_Click()
On Error GoTo Err_SaveRecord_Click

    DoCmd.RunCommand acCmdSaveRecord

Exit_SaveRecord_Click:
    Exit Sub

Err_SaveRecord_Click:
    MsgBox Err.Description
    Resume Exit_SaveRecord_Click
    
End Sub
