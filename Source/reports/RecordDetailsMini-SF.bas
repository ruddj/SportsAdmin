Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DateGrouping =1
    GridY =10
    Width =6462
    ItemSuffix =141
    Left =1500
    Top =240
    RecSrcDt = Begin
        0xe9d140bccedae140
    End
    RecordSource ="Report-Records"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x130b000037020000130b000096020000000000003e1900001801000001000000 ,
        0x020000007100000000000000a20700000100000000000000
    End
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin Chart
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            ControlSource ="ET_Des"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="ET_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Sex"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Age"
        End
        Begin BreakLevel
            ControlSource ="BestResult"
        End
        Begin PageHeader
            Height =0
            Name ="PageHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =1
            Name ="GroupHeader0"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =622
            BreakLevel =2
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    TextFontFamily =18
                    Top =56
                    Width =5391
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Field118"
                    ControlSource ="=\"EVENT: \" & [Sex Sub] & \" \" & [ET_Des]"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Top =396
                    Width =795
                    Height =225
                    FontWeight =700
                    Name ="Text126"
                    Caption ="DIVISION"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =851
                    Top =396
                    Width =1740
                    Height =225
                    FontWeight =700
                    Name ="Text127"
                    Caption ="NAME"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =2665
                    Top =396
                    Width =900
                    Height =225
                    FontWeight =700
                    Name ="Text128"
                    Caption ="TEAM"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3563
                    Top =396
                    Width =1080
                    Height =225
                    FontWeight =700
                    Name ="Text129"
                    Caption ="RESULT"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4656
                    Top =396
                    Width =735
                    Height =225
                    FontWeight =700
                    Name ="Text130"
                    Caption ="DATE"
                End
                Begin Line
                    BorderWidth =2
                    Top =396
                    Width =5430
                    Name ="Line140"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="AgeHeader"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =280
            OnFormat ="[Event Procedure]"
            Name ="Detail1"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    Left =850
                    Width =1761
                    Height =225
                    Name ="Field98"
                    ControlSource ="FullName"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =2721
                    Width =801
                    Height =225
                    TabIndex =1
                    Name ="Field117"
                    ControlSource ="CompetitorHouse"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =56
                    Width =681
                    Height =225
                    TabIndex =2
                    Name ="Field119"
                    ControlSource ="Age"

                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    Left =3628
                    Width =801
                    Height =225
                    TabIndex =3
                    Name ="Result"
                    ControlSource ="Result"

                End
                Begin TextBox
                    TextFontFamily =34
                    Left =4437
                    Width =156
                    Height =255
                    TabIndex =4
                    Name ="Field122"
                    ControlSource ="=LCase(Left([Units],1))"

                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    Left =4705
                    Width =606
                    Height =255
                    TabIndex =5
                    Name ="Field131"
                    ControlSource ="Date"

                End
                Begin Rectangle
                    BackStyle =0
                    Width =802
                    Height =280
                    Name ="Box135"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =802
                    Width =1852
                    Height =280
                    Name ="Box136"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =2649
                    Width =922
                    Height =280
                    Name ="Box137"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =3575
                    Width =1072
                    Height =280
                    Name ="Box138"
                End
                Begin Rectangle
                    BackStyle =0
                    Left =4652
                    Width =712
                    Height =280
                    Name ="Box139"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            BreakLevel =3
            OnFormat ="[Event Procedure]"
            Name ="AgeFooter"
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter2"
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

Dim AgeCount As Variant, PreviousResult As Variant

Private Sub AgeFooter_Format(Cancel As Integer, FormatCount As Integer)
  Rem AgeCount = 1
End Sub

Private Sub AgeHeader_Format(Cancel As Integer, FormatCount As Integer)

  AgeCount = 1
  
End Sub

Private Sub Detail1_Format(Cancel As Integer, FormatCount As Integer)
  
  If (AgeCount > 1) And (PreviousResult <> Me!Result) Then
    Cancel = True
  End If
  PreviousResult = Me!Result
  AgeCount = AgeCount + 1
  
End Sub
