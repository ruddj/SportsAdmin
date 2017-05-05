Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3344
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =-18930
    Top =5340
    Right =-15390
    Bottom =6480
    TimerInterval =1000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xef8ca22f10fae140
    End
    RecordSource ="SELECT [Misc-ReportsPopUp].* FROM [Misc-ReportsPopUp];"
    Caption ="Open Reports"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin Tab
            Width =5103
            Height =3402
            BorderLineStyle =0
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =1417
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =1199
                    Top =368
                    Width =891
                    BorderColor =12632256
                    Name ="ReportPopupX"
                    ControlSource ="ReportPopupX"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =234
                            Top =368
                            Width =915
                            Height =240
                            Name ="Label60"
                            Caption ="X Position:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =1198
                    Top =644
                    Width =891
                    TabIndex =1
                    BorderColor =12632256
                    Name ="ReportPopupY"
                    ControlSource ="ReportPopupY"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =233
                            Top =644
                            Width =915
                            Height =240
                            Name ="Label62"
                            Caption ="Y Position:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =113
                    Top =234
                    Width =2836
                    Height =731
                    Name ="Box65"
                End
                Begin CheckBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    Left =418
                    Top =150
                    TabIndex =2
                    BorderColor =12632256
                    Name ="Check63"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =1
                            OverlapFlags =247
                            Left =654
                            Top =113
                            Width =1500
                            Height =240
                            BackColor =-2147483633
                            Name ="Label64"
                            Caption ="Show Report Popup"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
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


Private Sub ReportPopupX_AfterUpdate()

  Call PositionForm
  
End Sub

Private Sub ReportPopupY_AfterUpdate()

  Call PositionForm

End Sub

Private Sub PositionForm()

On Error GoTo PositionForm_Exit
    
  Dim x As Integer, Y As Integer
  
  If VarEmpty(Me!ReportPopupX) Then
    Me!ReportPopupX = 0
  Else
    If Me!ReportPopupX < 0 Then
      Me!ReportPopupX = 0
    ElseIf Me!ReportPopupX > 11500 Then
      Me!ReportPopupX = 11500
    End If
  End If
  
  x = Me!ReportPopupX
    
  If VarEmpty(Me!ReportPopupY) Then
    Me!ReportPopupY = 0
  Else
    If Me!ReportPopupY < 0 Then
      Me!ReportPopupY = 0
    ElseIf Me!ReportPopupY > 9000 Then
      Me!ReportPopupY = 9000
    End If
  End If
  Y = Me!ReportPopupY
  
PositionForm_Exit:

End Sub
