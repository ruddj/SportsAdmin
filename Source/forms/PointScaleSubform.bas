Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    GridX =20
    GridY =20
    Width =3458
    ItemSuffix =18
    Left =4680
    Top =3045
    Right =6990
    Bottom =6885
    HelpContextId =50
    RecSrcDt = Begin
        0xbce33bd310cde140
    End
    RecordSource ="SELECT DISTINCTROW PointsScale.Place, PointsScale.Points, PointsScale.PtScale FR"
        "OM PointsScale ORDER BY PointsScale.Place, PointsScale.PtScale;"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
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
        Begin TextBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
        End
        Begin ListBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-236
        End
        Begin ComboBox
            LabelAlign =3
            BorderLineStyle =0
            Width =1701
            LabelX =-236
        End
        Begin FormHeader
            Height =72
            BackColor =-2147483633
            Name ="FormHeader1"
        End
        Begin Section
            Height =283
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =72
                    Width =675
                    Name ="Place"
                    ControlSource ="Place"
                    StatusBarText ="Place that is achieved by competitor"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =799
                    Width =795
                    TabIndex =1
                    Name ="Points"
                    ControlSource ="Points"
                    StatusBarText ="Points allocated to the competitor who gains this place"
                    FontName ="Tahoma"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1644
                    Width =336
                    Height =261
                    TabIndex =2
                    Name ="DeleteBut"
                    Caption ="Command17"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadadadadadaadad00adad00adaddadad00ad00adada ,
                        0xadadad0000adadaddadadad00adadadaadadad0000adadaddadad00ad00adada ,
                        0xadad00adad00adaddadadadadadadadaadadadadadadadaddadadadadadadada ,
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
                    ControlTipText ="Delete this entry."

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
Option Compare Database
Option Explicit

Private Sub DeleteBut_Click()
On Error GoTo Err_DeleteBut_Click
  Dim Response As Integer
  
  Response = MsgBox("Are you sure you want to delete this entry?", vbYesNo + vbInformation)
  If Response = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
  End If

Exit_DeleteBut_Click:
  DoCmd.SetWarnings True
  Exit Sub

Err_DeleteBut_Click:
    MsgBox Err.Description
    Resume Exit_DeleteBut_Click
    
End Sub
