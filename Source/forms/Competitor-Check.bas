Version =20
VersionRequired =20
Begin Form
    AllowAdditions = NotDefault
    Width =11436
    ItemSuffix =23
    Left =3510
    Right =11520
    Bottom =7590
    RecSrcDt = Begin
        0xc679aff38cdce140
    End
    RecordSource ="SELECT DISTINCTROW Competitors.PIN, Competitors.Gname, Competitors.Surname, Comp"
        "etitors.Sex, Competitors.H_Code, Competitors.H_ID, Competitors.DOB, Competitors."
        "Age FROM Competitors WHERE ((Competitors.Gname Is Null Or Competitors.Gname=\"\""
        ")) OR ((Competitors.Surname Is Null Or Competitors.Surname=\"\")) OR ((Competito"
        "rs.Sex Is Null Or Competitors.Sex=\"\")) OR ((Competitors.H_Code Is Null Or Comp"
        "etitors.H_Code=\"\")) OR ((Competitors.H_ID Is Null)) OR ((Competitors.DOB Is Nu"
        "ll)) OR ((Competitors.Age Is Null));"
    Caption ="Competitor Check"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
        End
        Begin OptionButton
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =187
            Height =187
        End
        Begin TextBox
            AutoLabel = NotDefault
            BorderLineStyle =0
        End
        Begin ListBox
            AutoLabel = NotDefault
            BorderLineStyle =0
        End
        Begin ComboBox
            AutoLabel = NotDefault
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =840
            BackColor =12632256
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Width =2730
                    Height =405
                    FontSize =14
                    Name ="Text6"
                    Caption ="Competitors Check"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1584
                    Top =600
                    Width =735
                    Height =240
                    Name ="Text8"
                    Caption ="Gname"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Top =600
                    Width =885
                    Height =240
                    Name ="Text10"
                    Caption ="Surname"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =3168
                    Top =600
                    Width =465
                    Height =240
                    Name ="Text12"
                    Caption ="Sex"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =3744
                    Top =600
                    Width =825
                    Height =240
                    Name ="Text14"
                    Caption ="H_Code"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =5040
                    Top =600
                    Width =540
                    Height =240
                    Name ="Text18"
                    Caption ="DOB"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    Left =6048
                    Top =600
                    Width =480
                    Height =240
                    Name ="Text20"
                    Caption ="Age"
                    FontName ="Tahoma"
                End
            End
        End
        Begin Section
            Height =240
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    Left =1584
                    Width =1575
                    ColumnWidth =1080
                    Name ="Gname"
                    ControlSource ="Gname"
                    StatusBarText ="Given Name(s)"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =87
                    Width =1590
                    TabIndex =1
                    Name ="Surname"
                    ControlSource ="Surname"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =93
                    Left =5040
                    Width =1020
                    ColumnWidth =945
                    TabIndex =2
                    Name ="DOB"
                    ControlSource ="DOB"
                    StatusBarText ="Date of Birth"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OverlapFlags =247
                    Left =6048
                    Width =780
                    ColumnWidth =840
                    TabIndex =3
                    Name ="Age"
                    ControlSource ="Age"
                    StatusBarText ="Age group (13 , 14, 15, OPEN, ETC.)"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =95
                    ListWidth =1690
                    Left =3168
                    Width =580
                    TabIndex =4
                    Name ="Field21"
                    ControlSource ="Sex"
                    RowSourceType ="Value List"
                    RowSource ="\"M\";\"F\""
                    ColumnWidths ="1440"
                    FontName ="Tahoma"

                End
                Begin ComboBox
                    OverlapFlags =247
                    ListWidth =1330
                    Left =3744
                    Width =1330
                    TabIndex =5
                    ColumnInfo ="\"\";\">\";\"10\";\"14\""
                    Name ="Field22"
                    ControlSource ="H_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="Select [H_Code] From [House];"
                    ColumnWidths ="1080"
                    FontName ="Tahoma"

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

Private Sub Form_Open(Cancel As Integer)

    Call MsgBox("Information for the following competitors is incomplete.  All fields need to have data in them.  The program may not perform as expected if fields are incomplete.", vbInformation)

End Sub
