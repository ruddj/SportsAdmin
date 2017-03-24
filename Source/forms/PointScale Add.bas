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
    GridX =5
    GridY =5
    Width =5329
    ItemSuffix =5
    Left =2040
    Top =165
    Right =11520
    Bottom =3465
    HelpContextId =50
    RecSrcDt = Begin
        0x2677765db6f2e140
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
                    InputMask =">CCCCCCCCCC"

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
                            Caption ="Enter the name of the new Scale:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4193
                    Top =906
                    Width =885
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="CreateBut"
                    Caption ="Create"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =451
                    Top =906
                    Width =915
                    Height =465
                    FontSize =8
                    FontWeight =400
                    TabIndex =2
                    Name ="CancelBut"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

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
On Error GoTo Err_CreateBut_Click

   ' Get New Event type Description and check

  If IsNull(DLookup("[PtScale]", "PointsScale", "[PtScale] = """ & [Description] & """")) Then

      If Me.OpenArgs = "RENAME" Then

        q1 = "UPDATE DISTINCTROW PointsScale SET PointsScale.PtScale = """ & [Description] & """"
        q1 = q1 & " WHERE PointsScale.PtScale=""" & Forms![PointScale]![PtScale] & """"

        q2 = "UPDATE DISTINCTROW Heats SET Heats.PtScale = """ & [Description] & """"
        q2 = q2 & " WHERE Heats.PtScale=""" & Forms![PointScale]![PtScale] & """"

        q3 = "UPDATE DISTINCTROW Final_Lev SET Final_Lev.PtScale = """ & [Description] & """"
        q3 = q3 & " WHERE Final_Lev.PtScale=""" & Forms![PointScale]![PtScale] & """"

        DoCmd.SetWarnings False

        DoCmd.RunSQL q1
        DoCmd.RunSQL q2
        DoCmd.RunSQL q3

        DoCmd.SetWarnings True

        DoCmd.Close
        
      Else

            Rept = DLookup("[R_Code]", "ReportTypes")
            Q = "INSERT INTO PointsScale ( PtScale, Place ) "
            Q = Q & "Values (""" & [Description] & """,0)"

            DoCmd.SetWarnings False
            DoCmd.RunSQL Q
            DoCmd.SetWarnings True
            
        DoCmd.Close

      End If

    Else
        MsgBox ("The event that you typed is already present.  Please choose a different event name.")
    End If

Exit_CreateBut_Click:
    
    DoCmd.SetWarnings True
    Exit Sub

Err_CreateBut_Click:
    MsgBox Error$
    Resume Exit_CreateBut_Click
    
End Sub

Private Sub Form_Load()

    If Me.OpenArgs = "RENAME" Then
        Me.[Caption] = "Rename PointScale"
        Me![CreateBut].Caption = "Rename"
    Else
        Me.[Caption] = "Add PointScale"
    End If
        
End Sub
