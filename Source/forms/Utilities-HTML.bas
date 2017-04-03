Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =7758
    ItemSuffix =140
    Left =2265
    Top =3285
    Right =9495
    Bottom =6975
    HelpContextId =540
    RecSrcDt = Begin
        0x8e1244042dc7e140
    End
    RecordSource ="MiscHTML"
    Caption ="Utilities 2"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    PrtMip = Begin
        0x8905000089050000890500008905000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000000000000
    End
    OnLoad ="[Event Procedure]"
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
        Begin CustomControl
            SpecialEffect =2
        End
        Begin FormHeader
            Height =0
            BackColor =128
            Name ="FormHeader1"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4882
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =144
                    Top =4176
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    HelpContextId =410
                    Name ="Help"
                    Caption ="Help"
                    OnClick ="Open Help"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6480
                    Top =4162
                    Width =1134
                    Height =510
                    FontSize =8
                    FontWeight =400
                    TabIndex =1
                    Name ="Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =2091
                    Top =432
                    Width =5106
                    TabIndex =2
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="Field123"
                    ControlSource ="ReportHeader"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =288
                            Top =432
                            Width =1740
                            Height =285
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text124"
                            Caption ="Web Page Header:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =2091
                    Top =837
                    Width =4536
                    TabIndex =3
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="TemplateFile"
                    ControlSource ="TemplateFile"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =288
                            Top =837
                            Width =1740
                            Height =285
                            FontWeight =400
                            BackColor =-2147483633
                            Name ="Text126"
                            Caption ="Template File:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6768
                    Top =792
                    Width =576
                    Height =351
                    FontWeight =400
                    TabIndex =4
                    Name ="TemplateBut"
                    Caption ="..."
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

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =2082
                    Top =1296
                    Width =4551
                    TabIndex =5
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="TemplateFileSummary"
                    ControlSource ="TemplateFileSummary"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =216
                            Top =1299
                            Width =1800
                            Height =285
                            FontWeight =400
                            Name ="Text129"
                            Caption ="Template File Summary:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6768
                    Top =1242
                    Width =576
                    Height =321
                    FontWeight =400
                    TabIndex =6
                    Name ="SummaryBut"
                    Caption ="..."
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
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    Left =2082
                    Top =1728
                    Width =4551
                    TabIndex =7
                    BackColor =16777215
                    BorderColor =12632256
                    Name ="HTMLlocation"
                    ControlSource ="HTMLlocation"
                    FontName ="Tahoma"

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =85
                            Left =210
                            Top =1725
                            Width =1800
                            Height =600
                            FontWeight =400
                            Name ="Text132"
                            Caption ="Location to put Generated Web Pages:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =6768
                    Top =1698
                    Width =576
                    Height =306
                    FontWeight =400
                    TabIndex =8
                    Name ="LocationBut"
                    Caption ="..."
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

Option Compare Database   'Use database order for string comparisons

Private Sub Close_Click()

    DoCmd.Close

End Sub

Private Sub LocationBut_Click()

Dim Result As String

On Error GoTo Err_LocationBut_Click

    Dim n As Variant

    n = BrowseFolder("Locate web folder") 'Me!ctlCommonDialog.FileName
    If n <> "" Then
        Me![HTMLlocation] = Trim(n)
    End If

Exit_LocationBut_Click:
    Exit Sub

Err_LocationBut_Click:
    MsgBox Error$
    Resume Exit_LocationBut_Click
    

End Sub

Private Sub SummaryBut_Click()

On Error GoTo Err_SummaryBut_Click

    Dim n As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    
    strFilter = ahtAddFilterItem(strFilter, "Web Files (*.htm)", "*.htm")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    n = ahtCommonFileOpenSave(InitialDir:="", _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate standard template file")
        
    If n <> "" Then
        Me![TemplateFileSummary] = Trim(n)
    End If

Exit_SummaryBut_Click:
    Exit Sub

Err_SummaryBut_Click:
    MsgBox Error$
    Resume Exit_SummaryBut_Click
    


End Sub

Private Sub TemplateBut_Click()
On Error GoTo Err_TemplateBut_Click

    Dim n As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    strFilter = ahtAddFilterItem(strFilter, "Web Files (*.htm)", "*.htm")
    strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
    
    n = ahtCommonFileOpenSave(InitialDir:="", _
        Filter:=strFilter, FilterIndex:=1, Flags:=lngFlags, _
        DialogTitle:="Locate standard template file")
        
    'Me!ctlCommonDialog.DialogTitle = "Locate standard template file"
    'Me!ctlCommonDialog.Filter = "Web Files (*.htm)|*.htm|All (*.*)|*.*"
    'Me!ctlCommonDialog.DefaultExt = "htm"
    'Me!ctlCommonDialog.FileName = ""
    
    'Me!ctlCommonDialog.ShowOpen
    
    'n = Me!ctlCommonDialog.FileName
    If n <> "" Then
        Me![TemplateFile] = Trim(n)
    End If

Exit_TemplateBut_Click:
    Exit Sub

Err_TemplateBut_Click:
    MsgBox Error$
    Resume Exit_TemplateBut_Click
    
End Sub
