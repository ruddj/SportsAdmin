Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7029
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =2070
    Top =1485
    Right =11925
    Bottom =7830
    HelpContextId =70
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x45821f397025e240
    End
    Caption ="Competitor - Quick Add"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
        Begin Section
            Height =4422
            BackColor =16767152
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    Left =422
                    Top =718
                    Width =2211
                    Height =256
                    Name ="Gname"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =427
                            Top =435
                            Width =2205
                            Height =256
                            Name ="Label1"
                            Caption ="First Name"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =422
                    Top =1318
                    Width =2211
                    Height =256
                    TabIndex =1
                    Name ="Surname"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =427
                            Top =1035
                            Width =2205
                            Height =256
                            Name ="Label3"
                            Caption ="Last Name"
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =87
                    ColumnCount =5
                    Left =3091
                    Top =400
                    Width =3799
                    Height =3817
                    TabIndex =7
                    Name ="ExisitngCompetitors"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;1985;455;852;287"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            Left =3090
                            Top =120
                            Width =2580
                            Height =240
                            FontWeight =700
                            BackColor =16767152
                            Name ="ListLabel"
                            Caption ="Competitors with similar name"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =285
                    Top =230
                    Width =2551
                    Height =3466
                    Name ="Box12"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    Left =377
                    Top =105
                    Width =1500
                    Height =225
                    FontWeight =700
                    BackColor =16767152
                    Name ="Label13"
                    Caption ="New Competitor"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    ColumnCount =3
                    Left =422
                    Top =2695
                    Width =2211
                    Height =256
                    TabIndex =4
                    ColumnInfo ="\"\";\">\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="H_code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT House.H_Code, House.H_NAme, House.H_ID FROM House WHERE (((House.Include)"
                        "=True)) ORDER BY House.H_NAme;"
                    ColumnWidths ="0;2155;0"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =427
                            Top =2412
                            Width =2205
                            Height =256
                            Name ="Label7"
                            Caption ="House"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    ColumnCount =2
                    ListWidth =1150
                    Left =416
                    Top =2008
                    Width =730
                    Height =256
                    TabIndex =2
                    Name ="Sex"
                    RowSourceType ="Value List"
                    RowSource ="\"M\";\"Male\";\"F\";\"Female\""
                    ColumnWidths ="271;631"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =407
                            Top =1725
                            Width =615
                            Height =256
                            BackColor =-2147483633
                            Name ="Text67"
                            Caption ="Gender"
                            FontName ="Arial"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    ListWidth =955
                    Left =1367
                    Top =2008
                    Width =1075
                    Height =256
                    TabIndex =3
                    HelpContextId =10000
                    ColumnInfo ="\"\";\"\";\"2\";\"1\""
                    Name ="Age"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Competitors.Age, IIf(IsNull([age]),Null,Val([age])) AS Expr1 FRO"
                        "M Competitors ORDER BY IIf(IsNull([age]),Null,Val([age]));"
                    ColumnWidths ="705"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =1366
                            Top =1725
                            Width =645
                            Height =256
                            BackColor =12632256
                            Name ="Text69"
                            Caption ="Age:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =247
                    Left =422
                    Top =3105
                    Width =2155
                    Height =397
                    FontWeight =700
                    TabIndex =5
                    Name ="AddCompetitorBut"
                    Caption ="Add Competitor"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =465
                    Top =3855
                    Width =2155
                    Height =397
                    TabIndex =6
                    Name ="CancelBut"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5880
                    Top =150
                    TabIndex =8
                    Name ="ShowAll"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="False"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6110
                            Top =120
                            Width =720
                            Height =240
                            Name ="Label17"
                            Caption ="Show all"
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
Option Compare Database
Option Explicit

Private Sub AddCompetitorBut_Click()
'On Error GoTo Err_AddCompetitorBut_Click

  If DataComplete Then
  
    Dim Crs As Recordset, Criteria As String
    
    Criteria = "[Surname]=""" & Me!Surname & """ AND [Gname]=""" & Me.Gname & """ "
    Criteria = Criteria & " AND [Sex]=""" & Me!Sex & """ AND [Age]=" & Me!Age
    Criteria = Criteria & " AND [H_code]=""" & Me!H_Code & """"
    
    Set Crs = CurrentDb.OpenRecordset("Competitors", dbOpenDynaset)
    
    Crs.FindFirst Criteria
    
    If Crs.NoMatch Then
      With Crs
        .AddNew
        !Surname = Me!Surname
        !Gname = Me!Gname
        !Sex = Me!Sex
        !Age = Me!Age
        !DOB = CDate("1/1/" & Year(Now) - Me!Age)
        !H_Code = Me!H_Code
        !H_ID = Me!H_Code.Column(2)
        .Update
      End With
            
      GlobalVariable = Me!Surname & ", " & Me!Gname
      GlobalCancel = False
      DoCmd.Close acForm, "Competitor-QuickAdd"
      
    Else
      MsgBox "There is another competitor with the same details.  You an not add two competitors with same details.", vbExclamation
    End If
      
    
  End If
  
Exit_AddCompetitorBut_Click:
    Exit Sub

Err_AddCompetitorBut_Click:
    MsgBox Err.Description
    Resume Exit_AddCompetitorBut_Click
    
End Sub
Private Sub CancelBut_Click()
On Error GoTo Err_CancelBut_Click

  GlobalCancel = True
  DoCmd.Close

Exit_CancelBut_Click:
    Exit Sub

Err_CancelBut_Click:
    MsgBox Err.Description
    Resume Exit_CancelBut_Click
    
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Err

  Const pLast = 1  ' processing last name
  Const pFirst = 2
  Const pAge = 3
  Const pSex = 4
  
  Dim n As String, i As Integer, c As String
  Dim LastName As String, FirstName As String, CurString As String
  Dim Action As Byte
  
  LastName = ""
  FirstName = ""
  CurString = ""
  Action = pLast
  n = Trim(Me.OpenArgs)
  i = 1
  
  Do Until i > Len(n)
    c = Mid(n, i, 1)
    
    If c = "," Then
      If Trim(CurString) <> "" Then
        Me!Surname = Trim(CurString)
      End If
      Action = pFirst
      CurString = ""
      
    ElseIf c = "|" Then
      If Trim(CurString) <> "" Then
        If Action = pFirst Then
          Me!Gname = Trim(CurString)
          Action = pAge
        ElseIf Action = pAge Then
          Me!Age = DLookup("[Cage]", "CompetitorEventAge", "[Eage]=""" & Trim(CurString) & """")
          Action = pSex
        Else ' Wouldn't normally get to here
          Me!Surname = Trim(CurString)
          Action = pAge
        End If
      End If
      
      CurString = ""
    
    Else
      CurString = CurString & c
    End If
    
    i = i + 1
    
  Loop
  
  If Action = pSex Then
    Me!Sex = Trim(CurString)
  End If
  
  Call SetCompetitorListRowSource(False)
  Dim Q As String
  
  Q = "SELECT DISTINCTROW Competitors.PIN, UCase([Surname]) & "", "" & [Gname] AS Expr1, Competitors.Age, House.H_Code, Competitors.Sex "
  Q = Q & "FROM House INNER JOIN Competitors ON House.H_Code = Competitors.H_Code "
  Q = Q & "WHERE [Gname] Like """ & Left(Me!Gname, 2) & "*"" And [Surname] Like """ & Left(Me!Surname, 2) & "*"""
  Q = Q & " ORDER BY [Surname], [Gname]"
  
  Me.ExisitngCompetitors.RowSource = Q
  

Form_Load_Exit:
  Exit Sub
  
Form_Load_Err:
  MsgBox "An error has occurred in [Form_Load]: " & Err.Description, vbCritical
  Resume Form_Load_Exit
  
End Sub

Private Function DataComplete() As Boolean
  
  If VarEmpty(Me!Gname) Or VarEmpty(Me!Surname) Or VarEmpty(Me!Sex) Or VarEmpty(Me!Age) Or VarEmpty(Me!H_Code) Then
    DataComplete = False
    MsgBox "All fields must be complete.", vbInformation
  Else
    DataComplete = True
  End If

End Function

Private Sub SetCompetitorListRowSource(AllCompetitors As Boolean)

  Dim Q As String
  
  Q = "SELECT DISTINCTROW Competitors.PIN, UCase([Surname]) & "", "" & [Gname] AS Expr1, Competitors.Age, House.H_Code, Competitors.Sex "
  Q = Q & "FROM House INNER JOIN Competitors ON House.H_Code = Competitors.H_Code "
  
  If Not AllCompetitors Then
    Q = Q & "WHERE [Gname] Like """ & Left(Nz(Me!Gname), 2) & "*"" And [Surname] Like """ & Left(Nz(Me!Surname), 2) & "*"""
    Me.ListLabel.Caption = "Competitors with similar name"
  Else
    Me.ListLabel.Caption = "All Competitors"
  End If
  
  Q = Q & " ORDER BY [Surname], [Gname]"
  
  Me.ExisitngCompetitors.RowSource = Q
  
End Sub


Private Sub Gname_AfterUpdate()

  If Not Me.ShowAll Then SetCompetitorListRowSource (False)

End Sub

Private Sub ShowAll_AfterUpdate()

  Call SetCompetitorListRowSource(Me.ShowAll)
  
End Sub

Private Sub Surname_AfterUpdate()

  If Not Me.ShowAll Then SetCompetitorListRowSource (False)

End Sub
