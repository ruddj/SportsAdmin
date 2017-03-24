Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =3968
    ItemSuffix =15
    Left =1380
    Top =540
    Right =6345
    Bottom =6945
    HelpContextId =520
    RecSrcDt = Begin
        0xd1dfe8b911cde140
    End
    RecordSource ="SELECT DISTINCTROW [Lane Promotion Allocation].ET_Code, [Lane Promotion Allocati"
        "on].Place, [Lane Promotion Allocation].Lane FROM [Lane Promotion Allocation] ORD"
        "ER BY [Lane Promotion Allocation].Place;"
    Caption ="Lane Allocation"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
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
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            Height =290
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =570
                    Height =230
                    Name ="Place"
                    ControlSource ="Place"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =720
                    Width =555
                    Height =230
                    TabIndex =1
                    Name ="Lane"
                    ControlSource ="Lane"
                    StatusBarText ="The lane allocated to the competitor who achieves Place."
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Arial"

                End
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2267
                    Width =786
                    TabIndex =2
                    Name ="Field14"
                    FontName ="Arial"

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

Private Sub Close_But_Click()
On Error GoTo Err_Close_But_Click


    DoCmd.Close

Exit_Close_But_Click:
    Exit Sub

Err_Close_But_Click:
    MsgBox Error$
    Resume Exit_Close_But_Click
    
End Sub

Private Sub Lane_BeforeUpdate(Cancel As Integer)

    If [Lane] = 0 Then
        MsgBox ("You can not have a lane 0.  Assign it another number (say 10).")
        Cancel = True

    End If
End Sub
