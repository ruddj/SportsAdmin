Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3263
    DatasheetFontHeight =11
    ItemSuffix =73
    Left =-11355
    Top =4890
    Right =-7830
    Bottom =12645
    HelpContextId =610
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x512d9814f1ede440
    End
    RecordSource ="tmpCEAM"
    Caption ="Event Ages to MeetManager Division Mappings"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            BorderLineStyle =0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =907
            BackColor =-2147483633
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =345
                    Top =285
                    Width =750
                    Height =555
                    FontSize =10
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="Label3"
                    Caption ="Event\015\012Age"
                    FontName ="Tahoma"
                    GroupTable =3
                    GridlineColor =10921638
                    LayoutCachedLeft =345
                    LayoutCachedTop =285
                    LayoutCachedWidth =1095
                    LayoutCachedHeight =840
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1155
                    Top =285
                    Width =1305
                    Height =555
                    FontSize =10
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="Label12"
                    Caption ="MeetManager\015\012Division #"
                    FontName ="Tahoma"
                    GroupTable =3
                    GridlineColor =10921638
                    LayoutCachedLeft =1155
                    LayoutCachedTop =285
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =840
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =3
                End
            End
        End
        Begin Section
            Height =396
            BackColor =-2147483633
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =330
                    Top =60
                    Width =750
                    Height =285
                    FontSize =10
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =-2147483617
                    Name ="Eage"
                    ControlSource ="Eage"
                    FontName ="Tahoma"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =330
                    LayoutCachedTop =60
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =345
                    DatasheetCaption ="Event Age"
                    BackTint =40.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1133
                    Top =56
                    Width =1305
                    Height =285
                    FontSize =10
                    TabIndex =1
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =-2147483617
                    Name ="Mdiv"
                    ControlSource ="Mdiv"
                    StatusBarText ="Meet Manager Division # Mapping"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Meet Manager Division # Mapping"
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =56
                    LayoutCachedWidth =2438
                    LayoutCachedHeight =341
                    DatasheetCaption ="MM Division"
                    ColumnStart =1
                    ColumnEnd =1
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin FormFooter
            Height =623
            BackColor =-2147483633
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =340
                    Width =1134
                    Height =510
                    FontSize =8
                    Name ="Close"
                    Caption ="&Done"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =340
                    LayoutCachedWidth =1474
                    LayoutCachedHeight =510
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1984
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    HelpContextId =50
                    Name ="HelpBut"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1984
                    LayoutCachedWidth =3118
                    LayoutCachedHeight =510
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
Option Compare Database
Option Explicit

Private Sub Close_Click()
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close
End Sub

Private Sub Form_Close()
    Dim tmpTable As String
    Dim strSQL As String
    
    tmpTable = "tmpCEAM"
    
    DoCmd.RunCommand acCmdSaveRecord
    
    ' Cleanup tmp table
    strSQL = "DELETE DISTINCTROW * FROM " & tmpTable
    CurrentDb.Execute strSQL, dbFailOnError
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim tmpTable As String
    Dim strSQL As String
    
    tmpTable = "tmpCEAM"

    ' Update tmp table with latest data
    strSQL = "DELETE DISTINCTROW * FROM " & tmpTable
    CurrentDb.Execute strSQL, dbFailOnError

    strSQL = "INSERT INTO " & tmpTable & " SELECT DISTINCT CompetitorEventAge.Eage, CompetitorEventAge.Mdiv FROM CompetitorEventAge "
    CurrentDb.Execute strSQL, dbFailOnError
    
    Me.Requery

End Sub

Private Sub Mdiv_AfterUpdate()
    Dim strSQL As String
    
    ' Update real data table after change made
    strSQL = "UPDATE CompetitorEventAge SET CompetitorEventAge.Mdiv = """ & Me.Mdiv & """ WHERE (CompetitorEventAge.Eage=""" & Me.Eage & """);  "
    CurrentDb.Execute strSQL, dbFailOnError
    
End Sub
