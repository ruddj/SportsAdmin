Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =20
    GridY =20
    Width =7511
    ItemSuffix =29
    Left =5475
    Top =2295
    Right =11520
    Bottom =7590
    RecSrcDt = Begin
        0x84f558081434e240
    End
    RecordSource ="SELECT DISTINCTROW Heats.E_Number, EventType.ET_Des, Events.Age, Events.Sex, Hea"
        "ts.F_Lev, Heats.Heat, Heats.E_Time FROM EventType INNER JOIN (Events INNER JOIN "
        "Heats ON Events.E_Code = Heats.E_Code) ON EventType.ET_Code = Events.ET_Code WHE"
        "RE (((EventType.Include)=Yes) AND ((Events.Include)=Yes)) ORDER BY Heats.E_Numbe"
        "r, EventType.ET_Des, Events.Age, Events.Sex DESC , Heats.F_Lev, Heats.Heat;"
    Caption ="Maintain Event Sequence"
    OnOpen ="[Event Procedure]"
    HelpFile ="SportsAdmin.chm"
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
        Begin ComboBox
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =0
            BackColor =12632256
            Name ="FormHeader1"
        End
        Begin Section
            Height =286
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =75
                    Top =15
                    Width =723
                    Height =256
                    Name ="E_Number"
                    ControlSource ="E_Number"
                    BeforeUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =828
                    Top =15
                    Width =1733
                    Height =256
                    TabIndex =1
                    Name ="Field4"
                    ControlSource ="ET_Des"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2598
                    Top =15
                    Width =573
                    Height =256
                    TabIndex =2
                    Name ="Field6"
                    ControlSource ="Sex"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3198
                    Top =15
                    Width =498
                    Height =256
                    TabIndex =3
                    Name ="Field8"
                    ControlSource ="Age"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =3723
                    Top =15
                    Width =423
                    Height =256
                    TabIndex =4
                    Name ="F_Lev"
                    ControlSource ="F_Lev"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4173
                    Top =15
                    Width =513
                    Height =256
                    TabIndex =5
                    Name ="Field27"
                    ControlSource ="Heat"
                    FontName ="Tahoma"

                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =4713
                    Top =15
                    Width =2073
                    Height =256
                    TabIndex =6
                    Name ="E_Time"
                    ControlSource ="E_Time"
                    StatusBarText ="Enter either the date, time or date and time."
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="Enter either the date, time or date and time."

                End
            End
        End
        Begin FormFooter
            Height =0
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

Dim In_Val As Variant
Dim TotRecs As Long
Dim LastTime As Variant

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

  If Me.Parent.SingleClickOption Then Call DetermineNewEventNum
  
End Sub

Private Sub E_Number_DblClick(Cancel As Integer)
  
  Call DetermineNewEventNum

End Sub

Private Sub E_Number_Enter()

    In_Val = Me![E_Number]

End Sub

Private Sub Field2_DblClick(Cancel As Integer)
  
  Dim Y
  
       Y = DMax("[E_Number]", "Heats") + 1

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


End Sub
