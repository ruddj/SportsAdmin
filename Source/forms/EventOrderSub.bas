Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6803
    DatasheetFontHeight =10
    ItemSuffix =1
    Left =6030
    Top =720
    Right =11520
    Bottom =4470
    HelpContextId =250
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb28f00a11c34e240
    End
    RecordSource ="SELECT DISTINCTROW Heats.E_Number, EventType.ET_Des, Events.Age, Events.Sex, Hea"
        "ts.F_Lev, Heats.Heat, Heats.E_Time FROM EventType INNER JOIN (Events INNER JOIN "
        "Heats ON Events.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.ET_Code WHE"
        "RE (((EventType.Include)=Yes) AND ((Events.Include)=Yes)) ORDER BY Heats.E_Numbe"
        "r, EventType.ET_Des, Events.Age, Events.Sex DESC , Heats.F_Lev, Heats.Heat;"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Line
            BorderLineStyle =0
            Width =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =298
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Width =723
                    Height =256
                    Name ="E_Number"
                    ControlSource ="E_Number"
                    BeforeUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =753
                    Width =1733
                    Height =256
                    TabIndex =1
                    Name ="Field4"
                    ControlSource ="ET_Des"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2523
                    Width =573
                    Height =256
                    TabIndex =2
                    Name ="Field6"
                    ControlSource ="Sex"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3123
                    Width =498
                    Height =256
                    TabIndex =3
                    Name ="Field8"
                    ControlSource ="Age"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3648
                    Width =423
                    Height =256
                    TabIndex =4
                    Name ="F_Lev"
                    ControlSource ="F_Lev"
                    FontName ="Arial"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4098
                    Width =513
                    Height =256
                    TabIndex =5
                    Name ="Field27"
                    ControlSource ="Heat"
                    FontName ="Arial"

                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4638
                    Width =2073
                    Height =256
                    TabIndex =6
                    Name ="E_Time"
                    ControlSource ="E_Time"
                    StatusBarText ="Enter either the date, time or date and time."
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Enter either the date, time or date and time."

                End
                Begin Line
                    LineSlant = NotDefault
                    OldBorderStyle =4
                    OverlapFlags =85
                    BorderLineStyle =3
                    Top =283
                    Width =6803
                    BorderColor =12632256
                    Name ="Line0"
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
Option Explicit

Private Sub E_Number_BeforeUpdate(Cancel As Integer)
    
  Dim new_val, ReOrder, i, Criteria, j
    
   TotRecs = DCount("*", "Heats")
   
   DoCmd.Hourglass True

   Dim tHeats As Recordset
   Set tHeats = CurrentDb.OpenRecordset("Heats", dbOpenDynaset)   ' Create dynaset.

   new_val = Me![E_Number]

   ReOrder = False
   
   If IsNull(new_val) Then
       ReOrder = True
   Else
       If Not IsNull(DLookup("[E_Number]", "Heats", "[E_Number]=" & new_val)) Then
           ReOrder = True
       End If
   End If

   If (new_val > TotRecs) Or (new_val < 1) Then
       Cancel = True
       MsgBox "The event number cannot exceed the total number of events or be less than 1!", 48
       ReOrder = False
   End If

   ' =======================================
   ' If Invalid event number
If ReOrder And (Me.Parent.AutoNumber = -1) Then


   ' ===========================================================================
   If (IsNull(new_val) And IsNull(In_Val)) Then
       'Do nothing.  This occurs when the user goes into a field, changes
       ' something but deletes it.


   ' ===========================================================================
   ElseIf (IsNull(new_val) And In_Val > 0) Then
       ' User is removing an event from the present ordering system.
       ' Slide all numbers after it up one to fill space

       For i = In_Val + 1 To TotRecs
           
           Criteria = "E_Number= " & i
           tHeats.FindFirst Criteria    ' Locate first occurrence."
       
           If Not tHeats.NoMatch Then
               tHeats.Edit
               tHeats!E_Number = tHeats!E_Number - 1
               tHeats.Update
           End If

       Next i
      

   ' ===========================================================================
   ElseIf new_val <> In_Val Or IsNull(In_Val) Then

     ' ===========================================================================
     ' Assuming that inserting before initial value.  ie changing event 10 to event 5
     ' Handles condition when In_Val = #null.  Sets In_Val to next unused event number -
     '   usually the last unused number.
     
     If (IsNull(In_Val) Or (In_Val > new_val)) Then

       If IsNull(In_Val) Then
           j = new_val + 1
           Criteria = "E_Number= " & j
           tHeats.FindFirst Criteria    ' Locate first occurrence."

           While Not tHeats.NoMatch
               j = j + 1
               Criteria = "E_Number= " & j
               tHeats.FindFirst Criteria    ' Locate first occurrence."
           Wend

           In_Val = j
       End If

       
       For i = In_Val - 1 To new_val Step -1
           
           Criteria = "E_Number= " & i
           tHeats.FindFirst Criteria    ' Locate first occurrence."
       
           If Not tHeats.NoMatch Then
               tHeats.Edit
               tHeats!E_Number = tHeats!E_Number + 1
               tHeats.Update
           End If

       Next i

    
     ' ===========================================================================
     ' Assuming that inserting after initial.  ie changing event 5 to event 10
      
     ElseIf In_Val < new_val Then


       For i = In_Val + 1 To new_val
           
           Criteria = "E_Number= " & i & ""
           tHeats.FindFirst Criteria    ' Locate first occurrence."
       
           If Not tHeats.NoMatch Then
               tHeats.Edit
               tHeats!E_Number = tHeats!E_Number - 1
               tHeats.Update
           End If

       Next i
      

     End If

   End If ' ================= ElseIf new_val <> In_Val Or IsNull(In_Val) Then

   tHeats.Close

  End If

  
Finish:

  DoCmd.Hourglass False
    
End Sub

Private Sub E_Number_Click()

  If Me.Parent!SingleClickOption.Value = True Then Call DetermineNewEventNum
  
End Sub

Private Sub E_Number_DblClick(Cancel As Integer)
  
  Call DetermineNewEventNum

End Sub

Private Sub E_Number_Enter()

    In_Val = Me!E_Number

End Sub

Private Sub Field2_DblClick(Cancel As Integer)
  
  Dim y
  
       y = DMax("[E_Number]", "Heats") + 1

End Sub

Private Sub E_time_AfterUpdate()

  LastTime = Me![E_Time]
  
End Sub

Private Sub E_time_Enter()
On Error GoTo Etime_Enter_Exit

  If Not VarEmpty(Me!E_Number) Then
    If IsNull(Me!E_Time) Then
      Me!E_Time = LastTime
    End If
  End If

Etime_Enter_Exit:

End Sub

Private Sub Form_Open(Cancel As Integer)

  LastTime = Null
  
End Sub

Private Sub DetermineNewEventNum()
Dim T As Long, i As Integer, NewEnum As Variant
 On Error GoTo DetermineNewEventNum_Err
 
  'MsgBox "DetermineNewEventNum"
  
    DoCmd.Hourglass True
    
    T = DCount("[Heat]", "Heats")

    NewEnum = DMax("[E_Number]", "Heats") + 1
    
    If NewEnum > T Then
        For i = 1 To T
            If IsNull(DLookup("[E_Number]", "Heats", "[E_Number]=" & i)) Then
                NewEnum = i
                GoTo ExitForLoop
            End If
        Next i
        
        GoTo Exit_DetermineNewEventNum

    End If

ExitForLoop:

    Me!E_Number = NewEnum

Exit_DetermineNewEventNum:
    
    DoCmd.Hourglass False

DetermineNewEventNum_Exit:
  Exit Sub
  
DetermineNewEventNum_Err:
  MsgBox Err.Description, vbExclamation
  Resume DetermineNewEventNum_Exit
  
End Sub
