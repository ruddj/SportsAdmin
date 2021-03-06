Option Compare Database   'Use database order for string comparisons
Option Explicit

Global GlobalFilename As Variant

Private Function EnsureDatabaseVersionIsCurrent(FileName) As Boolean
On Error GoTo EnsureDatabaseVersionIsCurrent_Err

  If SportsViewModule Then
    EnsureDatabaseVersionIsCurrent = False
    Exit Function
  End If

  Dim HasError As Boolean, db As Database
  
  Set db = DBEngine.Workspaces(0).OpenDatabase(FileName)
  
  SysCmd acSysCmdSetStatus, "Checking table: _AlwaysOpen"
  HasError = AddTable(FileName, "_AlwaysOpen")

  SysCmd acSysCmdSetStatus, "Checking table: CompetitorEventAge"
  HasError = AddTable(FileName, "CompetitorEventAge")
  
  SysCmd acSysCmdSetStatus, "Checking table: MiscHTML"
  HasError = AddTable(FileName, "MiscHTML")

  SysCmd acSysCmdSetStatus, "Applying field changes: 1 ... "
  HasError = AddField_nResult(db)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 2 ... "
  HasError = AddField_ProNum(db)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 3 ... "
  HasError = ChangeAgeFieldType(db)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 4 ... "
  HasError = AddField(db, "Competitors", "ID", dbText, False, 50)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 5 ... "
  HasError = AddField(db, "EventType", "PlacesAcrossAllHeats", dbBoolean, False, , False)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 6 ... "
  HasError = AddField(db, "Heats", "DontOverridePlaces", dbBoolean, False, , False)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 7 ... "
  HasError = AddField(db, "Heats", "EffectsRecords", dbBoolean, False, , False)
  
  SysCmd acSysCmdSetStatus, "Applying field changes: 8 ... "
  HasError = AddField(db, "Final_Lev", "EffectsRecords", dbBoolean, False, , False)
  
  SysCmd acSysCmdSetStatus, "Applying Meet Manager field changes: 1 ... "
  HasError = AddField(db, "Miscellaneous", "Mteam", dbText, False, 30)
  
  SysCmd acSysCmdSetStatus, "Applying Meet Manager field changes: 2 ... "
  HasError = AddField(db, "Miscellaneous", "Mcode", dbText, False, 4)
  
  SysCmd acSysCmdSetStatus, "Applying Meet Manager field changes: 3 ... "
  HasError = AddField(db, "Miscellaneous", "Mtop", dbInteger, False, 3, 3)
  
  SysCmd acSysCmdSetStatus, "Applying Meet Manager field changes: 4 ... "
  HasError = AddField(db, "EventType", "Mevent", dbText, False, 10)
  
  SysCmd acSysCmdSetStatus, "Applying Meet Manager field changes: 5 ... "
  HasError = AddField(db, "CompetitorEventAge", "Mdiv", dbText, False, 2)
  
  Set db = Nothing
  
  EnsureDatabaseVersionIsCurrent = HasError
  
EnsureDatabaseVersionIsCurrent_Exit:
  Exit Function
  
EnsureDatabaseVersionIsCurrent_Err:
  HasError = True
  GoTo EnsureDatabaseVersionIsCurrent_Exit
  
End Function

Function AddField(db As Database, tableName As String, _
                  fieldName As String, FieldType As Long, Required As Boolean, _
                  Optional FieldSize, Optional DefaultV)
  On Error Resume Next
  
  Dim td As TableDef
  Dim F As Field, Response As Variant
  Dim Q As String

  Set td = db.TableDefs(tableName)
  Set F = td.Fields(fieldName)
  
  If Err.Number <> 0 Then 'need to add field
    If IsMissing(FieldSize) Then
      Set F = td.CreateField(fieldName, FieldType)
    Else
      Set F = td.CreateField(fieldName, FieldType, FieldSize)
    End If
    
    F.Required = Required
    
    If Not IsMissing(DefaultV) Then F.DefaultValue = DefaultV
    td.Fields.Append F
    
    If fieldName = "EffectsRecords" Then
      If False And tableName = "Heats" Then
        Q = "A new field has been added to each heat that enables you to specify whether the heat should effect event records.  "
        Q = Q & "Use this feature to ensure that new records are set only for the events of your choosing, say the grand final races.  "
        MsgBox Q, vbInformation
      End If
      DoCmd.SetWarnings False
      DoCmd.RunSQL "UPDATE [" & tableName & "] SET [EffectsRecords]=TRUE"
      DoCmd.SetWarnings True
    End If
    
  End If
  
End Function

Function AddField_nResult(db As Database)

    On Error GoTo AddField_Err
    'Stop
    
    AddField_nResult = False

    'Dim CurrentDatabase As Database
    
    Dim td As TableDef
    Dim F As Field, Response As Variant
    Dim Indx As index
    

    ''*** Create nRecord field ****

    'Set db = DBEngine.Workspaces(0).OpenDatabase(FileName)
    'Set CurrentDatabase = DBEngine.Workspaces(0).Databases(0)
    
    Set td = db.TableDefs("Records")
    Set F = td.CreateField("nResult", dbSingle)
    td.Fields.Append F

    '*** Change type of Record field ****

    Response = MsgBox("To update to the latest version of the Sports Administrator it is necessary to remove all Event Records.  Do you wish to continue?", vbExclamation + vbYesNo + vbDefaultButton2, "Remove records")
    If Response = vbYes Then
        td.Fields.Delete "Result"    ' Delete field from collection.
        Set F = td.CreateField("Result", dbText)
    End If
    F.Size = 50
    td.Fields.Append F

    ' **** Remove old index and add new one ****
    ' I initially limited one record per event per day.  Bad.  Now there is no limitations

    td.Indexes.Delete "PrimaryKey" ' This was the name of the Original Index
    GoTo AddNewIndex

RemoveNewIndex:
    On Error Resume Next
    td.Indexes.Delete "PriIndex"
    On Error GoTo AddField_Err
    
AddNewIndex:
    Set Indx = td.CreateIndex("PriIndex")

    Indx.Primary = False
    Indx.Unique = False
    Set F = td.CreateField("E_Code", dbLong)
    Indx.Fields.Append F
    'Set F = TD.CreateField("Date", DB_DATE)
    'Indx.fields.Append F
    td.Indexes.Append Indx

    CurrentDb.Containers("Relationships").Documents.Refresh  ' Refresh possibly changed collection

AddField_Exit:

    Exit Function

AddField_Err:
    If Err = 3191 Then ' Field already exists
        ' Field has already been added to database which means that the Result field has been changed to Text data type
        ' Do nothing
        Resume RemoveNewIndex
    Else
        MsgBox ("An error has occured updating the RECORDS table.  The RECORD details may be inaccurate. Error:" & Error$)
        AddField_nResult = True
    End If

    GoTo AddField_Exit


End Function

Function AddField_ProNum(db)
On Error GoTo AddField_ProNum_Err

    'Stop
    
    AddField_ProNum = False ' No error

    'Dim db As Database, CurrentDatabase As Database
    
    Dim td As TableDef
    Dim F As Field, Response As Variant
    Dim Indx As index
    
    ''*** Create nRecord field ****

    'Set db = DBEngine.Workspaces(0).OpenDatabase(FileName)
    'Set CurrentDatabase = DBEngine.Workspaces(0).Databases(0)
    
    Set td = db.TableDefs("Final_Lev")
    Set F = td.CreateField("ProNum", dbInteger)
    td.Fields.Append F
    
AddField_ProNum_Exit:

    Exit Function

AddField_ProNum_Err:
    If Err = 3191 Then ' Field already exists
        ' Field has already been added to database which means that the Result field has been changed to Text data type
        ' Do nothing
    Else
        MsgBox ("An error has occured updating the FINAL_LEV table.  Error:" & Error$)
        AddField_ProNum = True ' An error has occured
    End If

    GoTo AddField_ProNum_Exit


End Function

Private Sub Test()

  Dim HasError As Boolean, db As Database
  
  Set db = DBEngine.Workspaces(0).OpenDatabase("e:\test.mdb")
  
  Call ChangeAgeFieldType(db)
  
End Sub

Function ChangeAgeFieldType(db As Database) As Boolean
On Error GoTo ChangeAgeFieldType_Err
  'Stop
  
  Dim td As TableDef, ErrorOccurred As Boolean, Q As String
  Dim F As Field, oF As Field, Response As Variant, i As index
  
  Set td = db.TableDefs("Competitors")
  Set oF = td.Fields("Age")
  
  If oF.type <> dbByte Then
    Q = "The age field for competitiors needs to be updated to the latest version.  "
    Q = Q & "You should ensure you have a backup of that carnival file before making these changes.  "
    Q = Q & "Do you wish to continue?"
    
    Response = MsgBox(Q, vbQuestion + vbYesNo + vbDefaultButton2)
    If Response = vbYes Then
      oF.Name = "AgeOld"
      oF.Required = False
      
      Set F = td.CreateField("Age", dbByte)
      F.DefaultValue = ""
      F.Required = True
      
      td.Fields.Append F
      
      Dim Rs As Recordset
      Set Rs = db.OpenRecordset("Competitors", dbOpenDynaset)
      Do Until Rs.BOF Or Rs.EOF
        Rs.Edit
        Rs!Age = Val(Rs!AgeOld)
        Rs.Update
        Rs.MoveNext
      Loop
      Rs.Close
      On Error Resume Next
      td.Indexes.Delete ("Age")
      td.Indexes.Delete ("Name&House")
      
      td.Fields.Delete ("AgeOld")
      
      Set i = td.CreateIndex("Age")
      i.Fields.Append i.CreateField("Age")
      td.Indexes.Append i
      
      Set i = td.CreateIndex("Name&House")
      i.Fields.Append i.CreateField("Surname")
      i.Fields.Append i.CreateField("Gname")
      i.Fields.Append i.CreateField("Age")
      i.Fields.Append i.CreateField("H_Code")
      i.Fields.Append i.CreateField("Sex")
      td.Indexes.Append i
      
      'Create some extra indexes while we are here
      Set i = td.CreateIndex("Surname")
      i.Fields.Append i.CreateField("Surname")
      td.Indexes.Append i
      
      Set i = td.CreateIndex("Gname")
      i.Fields.Append i.CreateField("Gname")
      td.Indexes.Append i
      
      
    Else
      ErrorOccurred = True
      GoTo ChangeAgeFieldType_Exit
    End If
  End If
  
  ErrorOccurred = False
  
ChangeAgeFieldType_Exit:
  ChangeAgeFieldType = ErrorOccurred
  Exit Function

ChangeAgeFieldType_Err:
  'Stop
  'Resume Next
  MsgBox "An error has occured updating the Age field in the Competitors table.  Error:" & Err.Description, vbCritical
  ErrorOccurred = True
  GoTo ChangeAgeFieldType_Exit

End Function


Function AddTable(FileName, NewTable)
    
    On Error Resume Next
    AddTable = False
    
    Dim Q As String
    
    Dim db As Database, CurrentDatabase As Database
    Dim td As TableDef, CurrentTD As TableDef
    Dim F As Field
    Set db = DBEngine.Workspaces(0).OpenDatabase(FileName)
    Set CurrentDatabase = CurrentDb()
    
    Set td = db.TableDefs(NewTable)
    'MsgBox ("Error1: Adding Table: " & Err)
    If Err = 3265 Then 'Table doesn't exist
        GoTo AddTable_Err
    ElseIf Err = 0 Then ' The table has already been added and ordered correctly.
        GoTo AddTable_Exit
    Else
        GoTo AddTable_Err2
    End If

AddTable_Exit:
    db.Close
    Set db = Nothing
    Set CurrentDatabase = Nothing
    Exit Function

AddTable_Err:

        'CurrentDatabase.TableDefs.Delete NewTable
        'MsgBox ("Error2: Adding Table: " & Err)
        If Err = 3265 Then
            ' The table doesn't exist which it should but does not cause a problem anyhow
        Else
            GoTo AddTable_Err2
        End If
        DoCmd.TransferDatabase acExport, "Microsoft Access", FileName, acTable, "zz~" & NewTable, NewTable, False
        'DoCmd.TransferDatabase acLink, "Microsoft Access", FileName, acTable, NewTable, NewTable, False

        GoTo AddTable_Exit

AddTable_Err2:
    MsgBox ("An error has occured adding the " & NewTable & " table.  The database integrity may be corrupt.")
    AddTable = True
    GoTo AddTable_Exit
    
End Function

Function AddTable_Competitors(FileName)

    AddTable_Competitors = False
    On Error Resume Next

    'Stop
    
    Dim Q As String
    
    Dim db As Database, CurrentDatabase As Database
    Dim td As TableDef, CurrentTD As TableDef
    Dim F As Field
    Set db = DBEngine.Workspaces(0).OpenDatabase(FileName)
    Set CurrentDatabase = CurrentDb()
    
    Set td = db.TableDefs("CompetitorsOrdered")
    If Err = 3265 Then 'Table doesn't exist
        GoTo AddTable_Competitors_Err
    ElseIf Err = 0 Then ' The table has already been added and ordered correctly.
        GoTo AddTable_Competitors_Exit
    Else
        GoTo AddTable_Competitors_Err2
    End If


AddTable_Competitors_Exit:
    db.Close
    Set db = Nothing
    Set CurrentDatabase = Nothing
    Exit Function

AddTable_Competitors_Err:

        CurrentDatabase.TableDefs.Delete "CompetitorsOrdered"
        If Err = 3265 Then
            ' The table doesn't exist which it should but doe not cause a p[roblem anyhow
        Else
            GoTo AddTable_Competitors_Err2
        End If
        DoCmd.TransferDatabase acExport, "Microsoft Access", FileName, acTable, "zz~CompetitorsOrdered", "CompetitorsOrdered", False
        DoCmd.TransferDatabase acLink, "Microsoft Access", FileName, acTable, "CompetitorsOrdered", "CompetitorsOrdered", False

        GoTo AddTable_Competitors_Exit

AddTable_Competitors_Err2:
    MsgBox ("An error has occured updating adding the COMPETITORS_ORDERED table.  The database integrity may be corrupt.")
    AddTable_Competitors = True
    GoTo AddTable_Competitors_Exit
    

End Function

Function Attach_Selected_File(ByVal IFID As Long, Posi As Variant, HasError As Variant) As Variant
'--------------------------------------------------------------------------------------------------------
' This function is used to attach all the tables for a selected INVENTORY file
' It determines the appropiate tables to attach from the table (Inventory Attached Tables)
' IFID is the code for the filename
' Posi is returned as the position in the variable GlobalFilename where the filename starts, before
' this position is the directory location
' HasError is returned true if an error was encountered
' The function only returns false if no file was selected form the dialog box
'
' Note: Use function [Attach_Selected_File2] if already know the filename for the new attachment

    On Error GoTo Err_Attach_Selected_File
    'Stop

    Dim Result As Variant, ReturnVal As Variant
    ReturnVal = True
    HasError = False
    '''Result = GetFileName("Select Database File", "Access Files (*.MDB)|*.MDB||", 1, "*")
    If (Result <> NoFileSelection) Then
        GlobalFilename = Trim$(Result)
        Result = Attach_Selected_File2(IFID, Posi, HasError, GlobalFilename)
    Else
        ReturnVal = False
    End If
Exit_Attach_Selected_File:
    Attach_Selected_File = ReturnVal
    Exit Function
Err_Attach_Selected_File:
    ReturnVal = False
    MsgBox Error$
    HasError = True
    Resume Exit_Attach_Selected_File
End Function

Function Attach_Selected_File2(ByVal IFID As Long, Posi As Variant, HasError As Variant, ByVal FileName As Variant) As Variant
'--------------------------------------------------------------------------------------------------------
' This function is used to attach all the tables for a selected INVENTORY file
' It determines the appropiate tables to attach from the table (Inventory Attached Tables)
' IFID is the code for the filename
' Posi is returned as the position in the variable GlobalFilename where the filename starts, before
' this position is the directory location
' HasError is returned true if an error was encountered
' The function only returns false if no file was selected form the dialog box
'
' Note: Use function [Attach_Selected_File] if you want to be given a file selection window

    On Error GoTo Err_Attach_Selected_File2
    Dim MyDb As Database, ITable As Recordset, SpecifiedPath As Variant, TT As TableDef, FTable As Recordset
    Dim DataExists As Variant, MyWS As Workspace, CPath  As Variant, AskUser  As Variant
    Dim Result As Variant, ReturnVal As Variant, db As Database, Rs As Recordset, Response As Variant
    ReturnVal = True
    HasError = False
    
    Set AlwaysOpenRS = Nothing
    
'Stop
    Posi = InStr(ReverseString(CStr(FileName)), "\")
    If Posi <> 0 Then
        
        ' Competitors Ordered in now local so don't need this
        'GlobalVariable = SysCmd(acSysCmdSetStatus, "Checking table: CompetitorsOrdered")
        'HasError = AddTable(FileName, "CompetitorsOrdered")
        Dim TableCount As Long
        
        TableCount = 0
        
        HasError = EnsureDatabaseVersionIsCurrent(FileName)
        
        Set MyWS = DBEngine.Workspaces(0)
        Set MyDb = CurrentDb

        'Call CheckRelationships(Filename)

        MyWS.BeginTrans
        On Error GoTo WKAttach_Selected_File2Error                                               ' Check Attachments
        Set ITable = MyDb.OpenRecordset("SELECT * FROM [Inventory Attached Tables] ORDER BY [Table Name]")
        Do Until ITable.EOF
          TableCount = TableCount + 1
            On Error Resume Next
            Set TT = MyDb.TableDefs(ITable![Table Name])
            If Err = 0 Then
                On Error GoTo WKAttach_Selected_File2Error
                TT.connect = ";DATABASE=" & CStr(FileName)
                SysCmd acSysCmdSetStatus, "Refreshing table: " & ITable![Table Name]
                TT.RefreshLink
            Else
                Set TT = MyDb.CreateTableDef(ITable![Table Name])
                On Error GoTo WKAttach_Selected_File2Error
                TT.connect = ";DATABASE=" & CStr(FileName)
                TT.SourceTableName = ITable![Table Name]
                GlobalVariable = SysCmd(acSysCmdSetStatus, "Attaching table: " & ITable![Table Name])
                MyDb.TableDefs.Append TT
                TT.RefreshLink
            End If
            ITable.MoveNext
            If TableCount = 1 Then ' Open this table and keep open to speed up linked table operations
              'MsgBox "Openign table"
              Call OpenAlwaysOpenRS
              'Set AlwaysOpenRS = CurrentDb.OpenRecordset(ITable![Table Name])
            End If
        Loop
        ITable.Close
        GlobalVariable = SysCmd(acSysCmdSetStatus, "Finalising settings ... ")
        MyWS.CommitTrans
        'GlobalVariable = SysCmd(acSysCmdSetStatus, "Ordering competitors ... ")
        'Call TransferToCompetitorOrdered
        On Error GoTo Err_Attach_Selected_File2
        MyWS.Close

        'GlobalVariable = SysCmd(acSysCmdSetStatus, "Applying field changes: 1 ... ")
        'HasError = AddField_nResult(FileName)   ' To update old carnival databases
        'GlobalVariable = SysCmd(acSysCmdSetStatus, "Applying field changes: 2 ... ")
        'HasError = AddField_ProNum(FileName)
        
    End If

Exit_Attach_Selected_File2:
    Attach_Selected_File2 = ReturnVal
    GlobalVariable = SysCmd(acSysCmdClearStatus)
    Exit Function

WKAttach_Selected_File2Error:
    MyWS.Rollback

Err_Attach_Selected_File2:
    ReturnVal = False
    MsgBox Error$, vbCritical
    HasError = True
    Resume Exit_Attach_Selected_File2

End Function

Function CheckInventoryAttached() As Variant
'---------------------------------------------------------------------------------------
'
    Set AlwaysOpenRS = Nothing
    
    On Error GoTo Err_CheckInventoryData
    Dim MyDb As Database, ITable As Recordset, SpecifiedPath As Variant, TT As TableDef, FTable As Recordset, TB  As TableDef
    Dim DataExists As Variant, MyWS As Workspace, CPath  As Variant, AskUser  As Variant, LTable As Recordset
    Dim DefaultLoc As Variant, DefaultLoc2 As Variant, Dummy As Variant, FileName As String, RPath As String, RFile As String
    Dim WhereCDF As Variant, RFPath As Variant, Response As Variant, FilePath As Variant

'    Stop
    Set MyWS = DBEngine.Workspaces(0)
    Set MyDb = CurrentDb()
    Set ITable = MyDb.OpenRecordset("Inventory Attached Tables", dbOpenDynaset)
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "UPDATE DISTINCTROW Carnivals SET Carnivals.Available = FileExists(GetCarnivalFullDir([Relative Directory]) & [Filename]);"
    DoCmd.SetWarnings True
    Set TB = MyDb.TableDefs("Competitors")
    FileName = UCase$(Right$(TB.connect, Len(TB.connect) - InStr(TB.connect, "=")))
    FilePath = Left$(FileName, Len(FileName) - InStr(ReverseString(FileName), "\") + 1)
    
    RPath = GetCarnivalRelDir(FilePath)
    RFPath = GetCarnivalFullDir(FileName)
    RFile = GetCarnivalFile(FileName)
    
    WhereCDF = "([Filename] = """ & RFile & """) AND ([Relative Directory] = """ & RPath & """)"
    
    If IsNull(DLookup("[Carnival]", "Carnivals", WhereCDF & " and [Available]")) Then
        DoCmd.OpenForm "Carnivals Maintain", A_NORMAL, , , , acDialog       ' then ask the user for their selection
        Call UpdateEventCompetitorAge
    Else
      Dim TableCount As Long
      TableCount = 0
      FileName = RFPath & RFile
      MyWS.BeginTrans
      On Error GoTo WKError                                               ' If all file locations ok then
      ITable.MoveFirst                                                    ' check tables available
      Do Until ITable.EOF
        TableCount = TableCount + 1
        On Error Resume Next
        Set TT = MyDb.TableDefs(ITable![Table Name])
        If Err = 0 Then
            On Error GoTo WKError
            TT.connect = ";DATABASE=" & FileName
            TT.RefreshLink
        Else
            On Error Resume Next
            Set TT = MyDb.CreateTableDef(ITable![Table Name])
            
            If Err.Number = 0 Or Err.Number = 3012 Then GoTo WKError ' 3012: Table already exists
            On Error GoTo WKError
            
            TT.connect = ";DATABASE=" & FileName
            TT.SourceTableName = ITable![Table Name]
            MyDb.TableDefs.Append TT
            TT.RefreshLink
            
            
        End If
        
        If TableCount = 1 Then ' Open this table and keep open to speed up linked table operations

          'Assumes that the _AlwaysOpen table has already been added
          Call OpenAlwaysOpenRS
          
        End If
        
        On Error GoTo WKError
        ITable.MoveNext
    
      Loop
      MyWS.CommitTrans
      On Error GoTo Err_CheckInventoryData
      ITable.Close                                                        ' If all inventory tables are available
    End If                                                                  ' then CHECK BILL DATA FILES
    
    CheckInventoryAttached = True
    
Exit_CheckInventoryData:
    Set MyDb = Nothing
    Exit Function
    
WKError:
    MyWS.Rollback
    
Err_CheckInventoryData:
    'MsgBox Error$ & "   Quitting Database."
    'DoCmd.Quit
    CheckInventoryAttached = False
    DoCmd.OpenForm "Carnivals Maintain", , , , , acDialog
    GoTo Exit_CheckInventoryData
End Function

Sub TestRelations()
  Call CheckRelationships("D:\Data\Sports\dist97\carnival\Demo-Sec Athletics.mdb")
End Sub
Sub CheckRelationships(FileName As Variant)

  If SportsViewModule Then Exit Sub
  
  Dim db As Database, WS As Workspace, NewDB As Database, Result As Variant
  Dim i As Integer, NR  As Relation, nF  As Field, r1 As Recordset, r2 As Recordset   ' Create Access Database
  Dim j As Integer, RelationError As Integer, RelationErrorNames As String
  Dim RelationName  As String
  
  On Error GoTo Err_CheckRelationships
  
  RelationErrorNames = ""

  Set WS = DBEngine.Workspaces(0)
  Set NewDB = WS.OpenDatabase(FileName)               ' Add relationships to
  Set db = WS.Databases(0)

  ' Check if all relationships are valid.  If not then delete all and recreate
  
  On Error GoTo Err_ValidatingRelationships
  
  RelationError = False
  Set r1 = db.OpenRecordset("zzz~Relationships Main", dbOpenSnapshot)       ' the database tables
  
  ' If the total number of relations are not correct then recreate all relationships
  If NewDB.Relations.Count <> DCount("[R ID]", "zzz~Relationships Main") Then
    RelationError = True
  End If

  ' Check each relation, its field and type.  If any inconsitencies recreate all relations
  ReturnVar = SysCmd(acSysCmdSetStatus, "Verifying relationships ... ")
  Do Until r1.EOF Or RelationError
    RelationName = r1![Relationship Name]
    Set NR = NewDB.Relations(RelationName)
    'If RelationName = "House-Competitors" Then Stop
    If NR.table <> r1![First Table] Or NR.foreignTable <> r1![Second Table] Or NR.Attributes <> r1![type] Then
      RelationError = True
    End If
    
    Set r2 = db.OpenRecordset("SELECT * FROM [zzz~Relationships Second] WHERE [R ID] = " & r1![R ID], dbOpenSnapshot, dbForwardOnly)
    Do Until r2.BOF Or r2.EOF Or RelationError
      Set nF = NR.Fields(r2![Field First])
      If nF.ForeignName <> r2![Field Second] Then RelationError = True
      r2.MoveNext
    Loop
    r1.MoveNext
  Loop
  
CreateNewRelationships: ' On relation problem exit to this point

  On Error GoTo Err_Creating_Relationships
  
  If RelationError Then
    If Debugging Then MsgBox ("Recreating relationships.")
    
    RelationError = False ' Reset the RelationErro flag
  
    On Error GoTo Err_Deleting_Relationships
    
    For j = (NewDB.Relations.Count - 1) To 0 Step -1
        NewDB.Relations.Delete NewDB.Relations(j).Name
    Next j
    On Error GoTo Err_Creating_Relationships

    r1.MoveFirst
    Do Until r1.EOF
      
        GlobalVariable = SysCmd(acSysCmdSetStatus, "Updating relationship " & r1![Relationship Name] & " ... ")
        Set NR = NewDB.CreateRelation(r1![Relationship Name])
        NR.table = r1![First Table]
        NR.foreignTable = r1![Second Table]
        NR.Attributes = r1![type]
        Set r2 = db.OpenRecordset("SELECT * FROM [zzz~Relationships Second] WHERE [R ID] = " & r1![R ID], dbOpenSnapshot, dbForwardOnly)
        Do Until r2.EOF
            Set nF = NR.CreateField(r2![Field First])
            nF.ForeignName = r2![Field Second]
            NR.Fields.Append nF
            r2.MoveNext
        Loop
        
        NewDB.Relations.Append NR
        r1.MoveNext
    Loop

  End If
    
  r1.Close
  r2.Close

Exit_Creating_Relationships:
  If RelationError Then
      MsgBox "There was a problem creating these relationships: " & RelationErrorNames & ".  This usually only occurs when accessing older carnivals.  This problem is not serious.  However small problems may arise.  Creating a new carnival file from scratch is the only way to resolve these issues.", vbInformation
  End If
  
  NewDB.Close
  Set NewDB = Nothing
  Set db = Nothing
  
  ReturnVar = SysCmd(acSysCmdClearStatus)
  Exit Sub

Err_CheckRelationships:
  MsgBox ("An unexpected error has occured in [CheckRelationships]: " & Err.Description)
  RelationError = True
  GoTo Exit_Creating_Relationships
  
Err_ValidatingRelationships: 'Error occured during relationship validation
  'MsgBox (Err.Number & ": " & Err.Description)
  RelationError = True
  GoTo CreateNewRelationships
  
Err_Deleting_Relationships:
  ' Assume Database is already open and relationships have been created
  GoTo Exit_Creating_Relationships

Err_Creating_Relationships:
  RelationError = True
  RelationErrorNames = RelationErrorNames & r1![Relationship Name] & " | "
  'MsgBox ("Error " & Error$ & ": creating relationship for: Table1=" & r1![First Table] & " Table2=" & r1![Second Table])
  Resume Next

End Sub

Function DBPath() As String
'-------------------------------------------------------------------------------------
' Returns the path to the current database or the error message
' Includes the final \

    On Error GoTo Err_DBPath
    Dim App As String, db As Database
    Set db = CurrentDb()
    App = db.Name
    DBPath = Left$(App, Len(App) - InStr(ReverseString(App), "\") + 1)
Exit_DBPath:
    Set db = Nothing
    Exit Function
    
Err_DBPath:
    DBPath = Error$
    Resume Exit_DBPath
End Function

Function ExtractDirectory(F As Variant) As Variant

    Dim Found As Boolean, X As Integer, L As Integer
    
    Found = False
    If IsNull(F) Then
        ExtractDirectory = ""
    Else
        L = Len(F)
        X = L
        'ExtractDirectory = F   ' was = Null
        ExtractDirectory = Null
    
        While Not Found And X > 0
            If Mid$(F, X, 1) = "\" Then
                Found = True
                ExtractDirectory = Left$(F, X)
            Else
                X = X - 1
            End If
        Wend
    End If
    
End Function

Function Make_File(ByVal FileName As String) As Variant
'--------------------------------------------------------------------------------
' Makes the file specified in the parameter and copies an empty image of all the tables
' that begin with "zz~". Removes the "zz~" when making the tables name.
'
' Returns TRUE upon successful completion
'

    On Error GoTo Err_Make_File
   
    Dim db As Database, WS As Workspace, NewDB As Database, Result As Variant                    ' Create Access Database
    Dim i As Integer, NR  As Relation, nF  As Field, r1 As Recordset, r2 As Recordset              ' and move in empty tables
    Dim iTrust As Integer, FilePath As String
            
    Result = False
    iTrust = -1
    Dim StrgSQL As String
    Set WS = DBEngine.Workspaces(0)
    DoCmd.SetWarnings False
    Set NewDB = WS.CreateDatabase(FileName, dbLangGeneral, dbVersion120)
    NewDB.Close
    Set db = CurrentDb()
    
    'MsgBox "If the file '" & fileName & "' is not in a trusted location, you will be prompted multiple times with an Access Security Notice. " & _
   ' "You will need to click multiple times on the Open button to continue.", vbInformation
    
    ' This is required as TransferDatabase will give many warning prompts if data file is not in Trusted Location
    FilePath = Left(FileName, InStrRev(FileName, "\") - 1)
    iTrust = AddTrustedLocation(FilePath, "Sports Admin Datafile", False)
    
    For i = db.TableDefs.Count - 1 To 0 Step -1
        If Left$(db.TableDefs(i).Name, 3) = "zz~" Then
            DoCmd.TransferDatabase acExport, "Microsoft Access", FileName, acTable, db.TableDefs(i).Name, _
           Right$(db.TableDefs(i).Name, Len(db.TableDefs(i).Name) - 3), False
        End If
    Next i
    
    Set NewDB = WS.OpenDatabase(FileName)                                                       ' Add relationships to
    Set r1 = db.OpenRecordset("zzz~Relationships Main", dbOpenSnapshot, dbForwardOnly)       ' the database tables
    Do Until r1.EOF
        Set NR = NewDB.CreateRelation(r1![Relationship Name])
        NR.table = r1![First Table]
        NR.foreignTable = r1![Second Table]
        NR.Attributes = r1![type]
        Set r2 = db.OpenRecordset("SELECT * FROM [zzz~Relationships Second] WHERE [R ID] = " & r1![R ID], dbOpenSnapshot, dbForwardOnly)
        Do Until r2.EOF
            Set nF = NR.CreateField(r2![Field First])
            nF.ForeignName = r2![Field Second]
            NR.Fields.Append nF
            r2.MoveNext
        Loop
        NewDB.Relations.Append NR
        r1.MoveNext
    Loop

    Result = True
Exit_Make_File:
    DelTrustedLocation (iTrust)
    Set db = Nothing
    DoCmd.SetWarnings True
    Make_File = Result
    Exit Function
    
Err_Make_File:
    MsgBox Error$
    Resume Exit_Make_File
End Function

Function NextCarnival() As String
'---------------------------------------------------------------------------------------
' Returns a new filename for carnival storage

    On Error GoTo Err_NextCarnival
    Dim i As Integer
    i = 1
    Do Until IsNull(DLookup("[Carnival]", "Carnivals", "[Filename] = ""CN" & String$(6 - Len(Trim$(CStr(i))), "0") & Trim$(CStr(i)) & ".ACCDB"""))
        i = i + 1
    Loop
    NextCarnival = "CN" & String$(6 - Len(Trim$(CStr(i))), "0") & Trim$(CStr(i)) & ".ACCDB"
Exit_NextCarnival:
    Exit Function
Err_NextCarnival:
    NextCarnival = Error$
    Resume Exit_NextCarnival
End Function

Function ReverseString(ByVal IStr As String) As String

    On Error GoTo Err_ReverseString
    Dim i As Variant, ReturnV As String
    ReturnV = ""
    For i = Len(IStr) To 1 Step -1
        ReturnV = ReturnV + Mid$(IStr, i, 1)
    
    Next
Exit_ReverseString:
    ReverseString = ReturnV
    Exit Function
Err_ReverseString:
    MsgBox Error$
    Resume Exit_ReverseString
End Function

Public Sub MaintainCompetitor(Action As String, PIN As Long)

On Error GoTo err_sdc

    Dim Criteria As String, db As Database, Crs As Recordset, CTrs As Recordset
    Dim NewTitle As String, Q As String
    
    ' Add competitor details to Competitor-Temp
    DoCmd.SetWarnings False
    DoCmd.RunSQL "delete * from [Competitors-Temp]"

    If Action = "EDIT" Then
        Q = "INSERT INTO [Competitors-Temp] ( PIN, Include, Gname, Surname, Sex, H_Code, H_ID, DOB, TotPts, Comments, Address1, Address2, Suburb, State, Postcode, Hphone, Wphone, Age, ID ) "
        Q = Q & "SELECT DISTINCTROW Competitors.PIN, Competitors.Include, Competitors.Gname, Competitors.Surname, Competitors.Sex, Competitors.H_Code, Competitors.H_ID, Competitors.DOB, Competitors.TotPts, Competitors.Comments, Competitors.Address1, Competitors.Address2, Competitors.Suburb, Competitors.State, Competitors.Postcode, Competitors.Hphone, Competitors.Wphone, Competitors.Age, Competitors.ID "
        Q = Q & "FROM Competitors WHERE Competitors.PIN=" & PIN
        DoCmd.RunSQL Q
    Else
        Q = "INSERT INTO [Competitors-Temp] ( Include, TotPts ) "
        Q = Q & "VALUES (true, 0)"
        DoCmd.RunSQL Q
    End If
    DoCmd.SetWarnings True

    GlobalCancel = False
    GlobalChange = False

    DoCmd.OpenForm "Competitors", , , , , acDialog, Action

    If Not GlobalCancel Then
      If GlobalChange Then
        Set db = CurrentDb()
        Set Crs = db.OpenRecordset("Competitors", dbOpenDynaset)   ' Create dynaset.
        Set CTrs = db.OpenRecordset("Competitors-Temp", dbOpenDynaset)   ' Create dynaset.
        
        Crs.FindFirst "[Pin]=" & PIN
        CTrs.MoveFirst
        If Action = "EDIT" Then
            Crs.Edit
        Else
            Crs.AddNew
        End If
        
        Crs!Include = CTrs!Include
        Crs!Gname = CTrs!Gname
        Crs!Surname = CTrs!Surname
        Crs!Sex = CTrs!Sex
        Crs!H_Code = CTrs!H_Code
        If Action = "ADD" Then
            Crs!H_ID = DLookup("[H_ID]", "House", "[H_Code]=""" & Crs!H_Code & """")
        Else
            Crs!H_ID = CTrs!H_ID
        End If
        Crs!DOB = CTrs!DOB
        Crs!TotPts = CTrs!TotPts
        Crs!Comments = CTrs!Comments
        Crs!Address1 = CTrs!Address1
        Crs!Address2 = CTrs!Address2
        Crs!Suburb = CTrs!Suburb
        Crs!State = CTrs!State
        Crs!Postcode = CTrs!Postcode
        Crs!Hphone = CTrs!Hphone
        Crs!Wphone = CTrs!Wphone
        Crs!Age = CTrs!Age
        Crs!ID = CTrs!ID
        
        Crs.Update
        
        'Call TransferToCompetitorOrdered
      End If

    End If
    
exit_sdc:
    Set db = Nothing
    Exit Sub

err_sdc:
    MsgBox (Error$)
    GoTo exit_sdc

End Sub

Sub TransferToCompetitorOrdered()

  Exit Sub
  

'On Error GoTo TransferToCompetitorOrdered_Err
  PleaseWaitMsg = "Updating Competitor Details ..."
  DoCmd.RunMacro "ShowPleaseWait"

  DoCmd.SetWarnings False
  DoCmd.RunSQL "UPDATE CompetitorsOrdered SET CompetitorsOrdered.Flag = No;"
  DoCmd.SetWarnings True
  
  Dim db As Database, Rs As Recordset, ors As Recordset, i As Integer, NoMoreRecords As Boolean
  Set db = CurrentDb
  
  Set Rs = db.OpenRecordset("CompetitorsOrderedQ", dbOpenSnapshot)
  Set ors = db.OpenRecordset("CompetitorsOrdered", dbOpenDynaset)
  i = 0
    
  If Rs.BOF Then GoTo TransferToCompetitorOrdered_Exit
  
  Do
    If ors.EOF Then
      ors.AddNew
    Else
      ors.Edit
    End If
    
'    If Left(ors!Surname, 6) = "aaaccc" Then Stop
    
    ors!PIN = Rs!PIN
    ors!Include = Rs!Include
    ors!Gname = Rs!Gname
    ors!Surname = Rs!Surname
    ors!Sex = Rs!Sex
    ors!H_Code = Rs!H_Code
    ors!H_ID = Rs!H_ID
    ors!TotPts = Rs!TotPts
    ors!Age = Rs!Age
    ors!Flag = True
    ors!Order = i
    ors.Update
    
    If Not ors.EOF Then ors.MoveNext
    
    Rs.MoveNext
    i = i + 1
    
  Loop Until Rs.EOF
  
  'DoCmd.SetWarnings False
  'DoCmd.RunSQL "DELETE DISTINCTROW CompetitorsOrdered.PIN FROM CompetitorsOrdered"
  'DoCmd.OpenQuery "Transfer Competitors to CompetitorsOrdered"
  'DoCmd.SetWarnings True

  Rs.Close
  ors.Close
  
  
TransferToCompetitorOrdered_Exit:
  DoCmd.RunMacro "ClosePleaseWait"

    Exit Sub

TransferToCompetitorOrdered_Err:
    MsgBox ("Error in Ordering Competitors.")
    DoCmd.RunMacro "ClosePleaseWait"
    GoTo TransferToCompetitorOrdered_Exit


End Sub

Public Sub OpenAlwaysOpenRS()
  
  SysCmd acSysCmdSetStatus, "Opening linked database..."
  
  Dim V As String
  
  On Error Resume Next
  V = AlwaysOpenRS.Name
  If Err.Number <> 0 Then 'Recordset is not open so open it
    Set AlwaysOpenRS = CurrentDb.OpenRecordset("_AlwaysOpen")
  End If
  
  SysCmd acSysCmdClearStatus
  
End Sub

Public Sub CloseAlwaysOpenRS()

  SysCmd acSysCmdSetStatus, "Closing linked database..."
  
  On Error Resume Next
  AlwaysOpenRS.Close
  
  SysCmd acSysCmdClearStatus
  
End Sub