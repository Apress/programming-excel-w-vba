VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetDBEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents mwb As Workbook
Attribute mwb.VB_VarHelpID = -1
Dim WithEvents mws As Worksheet
Attribute mws.VB_VarHelpID = -1

'SheetDBEngine Database Properties
'===================================================
Dim mstrLastRecord As String
Dim mstrdbRecordName As String
Dim mstrdbDataValidationList As String
Dim mstrdbSavedRecords As String
Dim mstrdbOneSide As String
Dim mstrdbManySide1 As String
Dim mstrdbManySide2 As String
Dim mstrdbManySide3 As String
Dim mstrdbManySide4 As String
Dim mstrdbManySidePrefix As String
Dim mstrdbManySideFirstColumn As String

Dim mintdbRecordsFirstRow As Integer
Dim mintdbOneSideColumnsCount As Integer
Dim mintdbManySideColumnsCount As Integer
Dim mintdbManySideRowsCount As Integer
Dim mintdbRangeOffset As Integer

'SheetDBEngine Public Properties
'===================================================
Dim mlngCalculation As Long
Dim mlngAbsolutePosition As Long


Public AutoSaveWorkbook As Boolean
Dim mbolSheetProtected As Boolean
Dim mbolScreenUpdating As Boolean
Dim mbolNewRecord As Boolean
Dim mbolDirty As Boolean
Dim mbolBOF As Boolean
Dim mbolEOF As Boolean

'SheetDBEngine Events
'===================================================
Event Current()
Event BeforeInsert(Cancel As Integer)
Event AfterInsert(Record As String)
Event BeforeUpdate(Cancel As Integer)
Event AfterUpdate(Record As String)
Event BeforeDelete(Cancel As Integer)
Event AfterDelete(Record As String)
Event BeforeSaveWorkbook(Cancel As Integer)

'SheetDBEngine Constant and Enumerators
'===================================================
Const mconDataBaseParameters = 15

Private Enum Operation
    LoadRecord = 1
    SaveRecord = 2
End Enum

Public Enum RecordPosition
    FirstRec = 1
    PreviousRec = -1
    NextRec = 2
    LastRec = 3
    NewRec = 0
End Enum

Private Sub Class_Initialize()
    Dim rg As Range
    
    Set mwb = ThisWorkbook
    Set mws = ActiveSheet
    
    Me.ScreenUpdating = True
    Me.Calculation = xlCalculationAutomatic
    mbolSheetProtected = mws.ProtectContents
    mstrdbDataValidationList = mws.Range("dbDataValidationList")
    mstrdbRecordName = mws.Range("dbRecordName")
    mstrdbSavedRecords = mws.Range("dbSavedRecords")
    mstrdbOneSide = mws.Range("dbOneSide")
    mstrdbManySide1 = mws.Range("dbManySide1")
    mstrdbManySide2 = mws.Range("dbManySide2")
    mstrdbManySide3 = mws.Range("dbManySide3")
    mstrdbManySide4 = mws.Range("dbManySide4")
    mstrdbManySideFirstColumn = mws.Range("dbManySideFirstColumn")
    mintdbRecordsFirstRow = mws.Range("dbRecordsFirstRow")
    mintdbOneSideColumnsCount = mws.Range("dbOneSideColumnsCount")
    mintdbManySideRowsCount = mws.Range("dbManySideRowsCount")
    mintdbManySideColumnsCount = mws.Range("dbManySideColumnsCount")
    mstrdbManySidePrefix = mws.Range("dbManySidePrefix")
    mintdbRangeOffset = mws.Range("dbRangeOffset")
    
    Set rg = mws.Range(mstrdbSavedRecords).Find(mws.Range(mstrdbDataValidationList), , , xlWhole)
    If Not rg Is Nothing Then
        Call SetAbsolutePosition(rg)
    End If
End Sub

Private Sub Class_Terminate()
    MsgBox "The SheetDBEngine Class for " & mws.Name & " sheet tab stop to run...", vbCritical, "Database Engine Fail"
End Sub

Public Property Let ScreenUpdating(Enabled As Boolean)
    mbolScreenUpdating = Enabled
    Application.ScreenUpdating = Enabled
End Property

Public Property Get ScreenUpdating() As Boolean
    ScreenUpdating = mbolScreenUpdating
End Property

Public Property Let Calculation(CalculateMethod As XlCalculation)
    mlngCalculation = CalculateMethod
    Application.Calculation = xlCalculationAutomatic
End Property

Public Property Get Calculation() As XlCalculation
    Calculation = mlngCalculation
End Property

Private Sub SetAbsolutePosition(rg As Range)
    mlngAbsolutePosition = (rg.Row - mintdbRecordsFirstRow)
End Sub

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = mlngAbsolutePosition
End Property

Public Property Get BOF() As Boolean
    BOF = mbolBOF
End Property

Public Property Get Dirty() As Boolean
    Dirty = mbolDirty
End Property
 
Public Property Get EOF() As Boolean
    EOF = mbolEOF
End Property

Public Property Get NewRecord() As Boolean
    NewRecord = mbolNewRecord
End Property

Public Property Get RecordCount() As Long
    RecordCount = mws.Range(mstrdbSavedRecords).Rows.Count - 1
End Property

Private Sub mwb_BeforeClose(Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    Dim strRecord  As String
    Dim bolSaved As Boolean

    If Me.Dirty Then
        strRecord = mws.Range(mstrdbDataValidationList)
        strTitle = "Save " & mstrdbRecordName & " data?"
        If Len(strRecord) = 0 Then
            strMsg = "New " & mstrdbRecordName & " data had been changed." & vbCrLf
        Else
            strMsg = mstrdbRecordName & "" & strRecord & "' data had been changed." & vbCrLf
        End If
        strMsg = strMsg & "Save data before close the workbook?"
        Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion, strTitle)
            Case vbYes
                bolSaved = Me.SaveAs(strRecord)
                Cancel = Not bolSaved
            Case vbNo
                If Len(strRecord) = 0 Then
                    Application.EnableEvents = False
                    Call Clear
                    mws.Range(mstrdbDataValidationList) = "New " & mstrdbRecordName
                    mws.Range(mstrdbDataValidationList).Select
                End If
            Case vbCancel
                Cancel = True
        End Select
    End If
End Sub

Private Sub mWs_Change(ByVal Target As Range)
    Select Case Target.Address
        Case Is = mws.Range(mstrdbDataValidationList).Address
            'User is trying to load a New Record
            TryToLoadSelectedRecord
        Case Else
            'Sheet data has changed
            mbolDirty = True
            If mws.Range(mstrdbDataValidationList) = "New " & mstrdbRecordName Then
                Application.EnableEvents = False
                    mws.Range(mstrdbDataValidationList) = ""
                Application.EnableEvents = True
            End If
     End Select
End Sub

Private Sub mWs_SelectionChange(ByVal Target As Range)
    If mws.Range(mstrdbDataValidationList) = "" Then
        mstrLastRecord = "New " & mstrdbRecordName
    Else
        mstrLastRecord = mws.Range(mstrdbDataValidationList)
    End If
End Sub

Private Sub Echo(fEnable As Boolean)
    With Application
        .ScreenUpdating = (fEnable And Me.ScreenUpdating)
        .EnableEvents = fEnable
        .Calculation = IIf(fEnable, Me.Calculation, xlManual)
    End With
End Sub

Private Sub TryToLoadSelectedRecord()
    Dim strMsg As String
    Dim strNewRecord As String
    
    strNewRecord = mws.Range(mstrdbDataValidationList)
    
    'Verify if current Record had been changed
    If Me.Dirty Then
        'Save current Record before change it?
        strMsg = mstrLastRecord & " data had been changed." & vbCrLf & vbCrLf
        strMsg = strMsg & "Save " & mstrLastRecord & " before load '" & strNewRecord & "'?"
        If MsgBox(strMsg, vbQuestion + vbYesNo, "Save current data?") = vbYes Then
            If Not SaveAs(mstrLastRecord) Then
                'Record data not saved!
                Application.EnableEvents = False
                    mws.Range(mstrdbDataValidationList) = mstrLastRecord
                Application.EnableEvents = True
                Exit Sub
            End If
        End If
        mbolDirty = False
    End If
        
    'Load selected Record data
    Call Load(strNewRecord)
End Sub

Public Sub Load(Optional strRecord As String)
    'Disable screen updating, events and recalc
    Call Echo(False)
        Select Case strRecord
            Case "", "New " & mstrdbRecordName
                'User selected a "New Record"
                Call Clear
                mws.Range(mstrdbDataValidationList) = "New " & mstrdbRecordName
                mbolNewRecord = True
                'Set record position
                mbolBOF = True
                mbolEOF = True
                mlngAbsolutePosition = 0
            Case Else
                'User selected a saved Record. Load it!
                Call LoadSaveData(strRecord, LoadRecord)
                
                On Error Resume Next
                mws.Range(mstrdbDataValidationList).Select
                mstrLastRecord = strRecord
                mbolNewRecord = False
        End Select
        mbolDirty = False
    'Enabledscreen updating, events and recalc
    Call Echo(True)
    
    'Raise Current event
    RaiseEvent Current
End Sub

Private Sub LoadSaveData(strRecord As String, Perform As Operation)
    Dim rg As Range
    Dim rgCells As Range
    Dim rgArea As Range
    Dim rgAreaColumn As Range
    Dim strRangeName As String
    Dim strRelation As String
    Dim intOffSet As Integer
    Dim intRelation As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intAreaCol As Integer
    Dim intMaxRows As Integer

    Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord, , , xlWhole)
    'Set record position
    Call SetAbsolutePosition(rg)
    mbolBOF = (rg.Row = mws.Range(mstrdbSavedRecords).Row + 1)
    mbolEOF = (rg.Row = mws.Range(mstrdbSavedRecords).Row + mws.Range(mstrdbSavedRecords).Rows.Count - 1)
    
    'Load/Save one side worksheet records (one cell at a time)
    If Len(mstrdbOneSide) Then
        Set rgCells = mws.Range(mstrdbOneSide)
        For Each rgArea In rgCells.Areas
            For intRow = 1 To rgArea.Rows.Count
                For intCol = 1 To rgArea.Columns.Count
                    If Perform = SaveRecord Then
                        rg.Offset(0, mintdbRangeOffset + intOffSet) = rgArea.Cells(intRow, intCol)
                    Else
                        rgArea.Cells(intRow, intCol) = rg.Offset(0, mintdbRangeOffset + intOffSet)
                    End If
                    intOffSet = intOffSet + 1
                    If rgArea.Cells(intRow, intCol).MergeCells Then
                        intRow = intRow + rgArea.Cells(intRow, intCol).MergeArea.Rows.Count - 1
                        intCol = intCol + rgArea.Cells(intRow, intCol).MergeArea.Columns.Count - 1
                    End If
                Next
            Next
        Next
    End If
    
    'Load/Save many side worksheet records
    strRangeName = mstrdbManySidePrefix & FixName(strRecord)
    'Process each many-side records range Relation
    intRow = 0
    For intRelation = 1 To 4
        strRelation = Choose(intRelation, mstrdbManySide1, mstrdbManySide2, mstrdbManySide3, mstrdbManySide4)
        If Len(strRelation) Then
            intCol = 0
            intMaxRows = 0
            Set rgCells = mws.Range(strRelation)
            For Each rgArea In rgCells.Areas
                For intAreaCol = 0 To rgArea.Columns.Count - 1
                    Set rg = mws.Range(strRangeName).Offset(intRow, intCol)
                    Set rg = rg.Resize(rgArea.Rows.Count, 1)
                    Set rgAreaColumn = mws.Range(mws.Cells(rgArea.Row, rgArea.Column + intAreaCol), _
                                                 mws.Cells(rgArea.Row + rgArea.Rows.Count - 1, rgArea.Column + intAreaCol))
                    If Perform = SaveRecord Then
                        rg.Value = rgAreaColumn.Value
                    Else
                        rgAreaColumn.Value = rg.Value
                    End If
                    
                    If rgArea.Cells(1, intAreaCol + 1).MergeCells Then
                        intAreaCol = intAreaCol + rgArea.Cells(1, intAreaCol + 1).MergeArea.Columns.Count - 1
                    End If
                    intCol = intCol + 1
                Next
                
                If intMaxRows < rgArea.Rows.Count Then
                    intMaxRows = rgArea.Rows.Count
                End If
            Next
            intRow = intRow + intMaxRows + 1
        End If
    Next
End Sub

Public Function SaveAs(Optional strLastRecord As String) As Boolean
    Dim strRecord As String
    Dim bolNewRecord As Boolean
    
    'Verify if Record data is still empty
    strRecord = mws.Range(mstrdbDataValidationList)
    If Not Me.Dirty And strRecord = "New " & mstrdbRecordName Then
        Exit Function
    End If

    If strLastRecord = "" Then
        strLastRecord = strRecord
    End If
    strRecord = GetRecordName(strLastRecord, bolNewRecord)
    
    If Len(strRecord) Then
        SaveAs = Save(strRecord, bolNewRecord)
   End If
End Function

Public Function Save(strRecord As String, Optional bolNewRecord As Boolean) As Boolean
    Dim intCancelInsert As Integer
    Dim intCancelUpdate As Integer
    Dim intCancelSave As Integer
    Dim intCancelSaveWorkbook As Integer
    Dim bolRecordSaved As Boolean
    
    'Raise events BeforeInsert and BeforeUpdate (allow cancel operation)
    If bolNewRecord Then
        RaiseEvent BeforeInsert(intCancelInsert)
        If intCancelInsert Then
            Exit Function
        End If
    End If
    
    RaiseEvent BeforeUpdate(intCancelUpdate)
    If intCancelUpdate Then
        Exit Function
    End If
        
    'Disable application events to allow cell change by macro code
    Call Echo(False)
        mws.Unprotect
            bolRecordSaved = SaveData(strRecord, bolNewRecord)
        If mbolSheetProtected Then
            mws.Protect
        End If
        
        If bolRecordSaved Then
            'Update record properties
            mbolNewRecord = False
            mbolDirty = False
            
            'Define current Record as saved Record
            On Error Resume Next
            mws.Range(mstrdbDataValidationList) = strRecord
            mws.Range(mstrdbDataValidationList).Select
    
            'Raise events AfterUpdate. AfterInsert, Current and BeforeSaveWorkbook
            RaiseEvent AfterUpdate(strRecord)
            If bolNewRecord Then
                RaiseEvent AfterInsert(strRecord)
            End If
            RaiseEvent Current
            
            'Save the worbook after save the record?
            If Me.AutoSaveWorkbook Then
                RaiseEvent BeforeSaveWorkbook(intCancelSaveWorkbook)
                If Not intCancelSaveWorkbook Then
                    ActiveWorkbook.Save
                End If
            End If
            
            Save = True
        Else
            MsgBox "There is no more room to save data on this worksheet!", vbCritical, "Can't save data"
        End If
    Call Echo(True)
End Function

Private Function SaveData(strRecord As String, Optional bolNewRecord As Boolean) As Boolean
    Dim rg As Range
    Dim rgData As Range
    Dim strRangeName As String
    Dim lngRow As Long
    Dim bolWorksheetIsFull As Boolean
    
    Set rg = mws.Range(mstrdbSavedRecords)
    If bolNewRecord Then
        'Define sheet row where next Record data will be stored
        lngRow = NextEntryRow(bolWorksheetIsFull)
        
        'Verify if sheet is full
        If bolWorksheetIsFull Then
            'No more room to save data
            Exit Function
        End If
        
        'Verify if mstrSavedRecords last rows is a empty cell
        If Not rg.Cells(rg.Rows.Count, 1) = "" Then
            'Insert a new row at bottom of SavedRecords range name and update rg object
            rg.Resize(rg.Rows.Count + 1).Name = "'" & mws.Name & "'!" & mstrdbSavedRecords
            Set rg = mws.Range(mstrdbSavedRecords)
        End If

        'Position on new cell of SavedRecords range and save New Record name
        rg.Cells(rg.Rows.Count, 1) = strRecord
        
        If Len(mstrdbManySide1) Then
            'Define Record name as 'rec_<strRecord>' and create it range name
            strRangeName = mstrdbManySidePrefix & FixName(strRecord)
            mws.Names.Add strRangeName, "='" & mws.Name & "'!" & mstrdbManySideFirstColumn & lngRow
            mws.Names(strRangeName).Visible = False
        End If
    End If
    
    Call LoadSaveData(strRecord, SaveRecord)
        
    'Sort SavedRecords range keeping "New Record" on the top of the list
    Set rgData = mws.Range(mws.Cells(rg.Row + 1, rg.Column), _
                           mws.Cells(rg.Row + rg.Rows.Count, rg.Column + mintdbRangeOffset + mintdbOneSideColumnsCount - 1))
    rgData.EntireRow.Hidden = False
        rgData.Sort rg.Cells(, 1)
        Set rg = rgData.Find(strRecord, , , xlWhole)
        Call SetAbsolutePosition(rg)
    rgData.EntireRow.Hidden = True

    On Error Resume Next
    mws.Range("A1").Select
    SaveData = True
End Function

Private Function GetRecordName(strRecord As String, bolNewRecord As Boolean) As String
    Dim rg As Range
    Dim strNewRecord As String
    Static sintDefaultName As Integer
    
    If strRecord = "" Then
        sintDefaultName = sintDefaultName + 1
        strRecord = "New " & mstrdbRecordName & Replace(Date, "/", "_")
        If sintDefaultName > 1 Then
            strRecord = strRecord & " " & sintDefaultName
        End If
        strRecord = InputBox("Data will be saved as:", "Confirm data name", strRecord)
    End If
    
    If Len(strRecord) Then
        'Verify if strRecord already exist on mstrdbSaveRecords
        Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord)
    
        If rg Is Nothing Then
            bolNewRecord = True
        Else
            'Confirm proposed record name
            strNewRecord = InputBox(mstrdbRecordName & " '" & strRecord & "' already exist. Do you want to overwrite it?", _
                                    "Overwrite " & strRecord & " data?", strRecord)
            'Verify if record name changed
            If strRecord <> strNewRecord Then
                'Proposed record name changed. Verify if new name alteady exist
                Set rg = mws.Range(mstrdbSavedRecords).Find(strNewRecord)
                If rg Is Nothing Then
                    bolNewRecord = True
                Else
                    'New name already exist. Confirm overwrite
                    If MsgBox("The name you typed, '" & strNewRecord & "', already exist. Overwrite it?", _
                              vbYesNo + vbDefaultButton2 + vbQuestion, _
                              "Overwrite '" & strNewRecord & "'?") = vbNo Then
                        strNewRecord = ""
                    End If
                End If
                strRecord = strNewRecord
            End If
        End If
    End If
        
    GetRecordName = strRecord
End Function

Private Function NextEntryRow(bolWorksheetIsFull As Boolean) As Long
    Dim lngRow As Long

    If Len(mstrdbManySide1) Then
        'Use many-side records to find next entry row
        lngRow = mintdbRecordsFirstRow + (mws.Range(mstrdbSavedRecords).Rows.Count - 1) * mintdbManySideRowsCount
        If lngRow < mws.UsedRange.Rows.Count Then
            lngRow = mintdbRecordsFirstRow + (mws.Range(mstrdbSavedRecords).Rows.Count) * mintdbManySideRowsCount
        End If
    Else
        'Just one-side record to find next entry row
        lngRow = mintdbRecordsFirstRow + mws.Range(mstrdbSavedRecords).Rows.Count
        bolWorksheetIsFull = (lngRow > (ActiveSheet.Rows.Count - mws.Range(mstrdbSavedRecords).Rows.Count))
    End If
    NextEntryRow = lngRow
End Function

Private Sub Clear()
    Dim rgCells As Range
    Dim strRange As String
    Dim intI As Integer
    
    'Clear one side worksheet records
    If Len(mstrdbOneSide) Then
        Set rgCells = mws.Range(mstrdbOneSide)
        rgCells = ""
    End If

    'Clear many side worksheet records
    For intI = 1 To 4
        strRange = Choose(intI, mstrdbManySide1, mstrdbManySide2, mstrdbManySide3, mstrdbManySide4)
        If Len(strRange) Then
            Set rgCells = mws.Range(strRange)
            rgCells = ""
        End If
    Next
End Sub

Public Sub Delete()
    Dim strRecord As String
    Dim strMsg As String
    Dim strTitle As String
    Dim bolNewRecord As Boolean
    
    strRecord = mws.Range(mstrdbDataValidationList)
    If strRecord = "" Or strRecord = "New " & mstrdbRecordName Then
        If mbolDirty Then
            bolNewRecord = True
            strMsg = "New " & mstrdbRecordName & " data has not been saved yet." & vbCrLf
            strMsg = strMsg & "Do you want to delete it?"
            strTitle = "Delete unsaved record?"
        Else
            Exit Sub
        End If
    Else
        strMsg = "Do you want to delete " & strRecord & " record?"
        strTitle = "Delete record?"
    End If
        
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion, strTitle) = vbYes Then
        If DeleteRecord(strRecord, bolNewRecord) Then
            'Define current Record as New Record
            mstrLastRecord = "New " & mstrdbRecordName
            mws.Range(mstrdbDataValidationList) = mstrLastRecord
        End If
    End If
End Sub

Public Function DeleteRecord(strRecord As String, Optional NewRecord As Boolean) As Boolean
    Dim intCancelDelete As Integer
    Dim intCancelSaveWorkbook As Integer
        
    'Raise event BeforeDelete
    RaiseEvent BeforeDelete(intCancelDelete)
    If intCancelDelete Then
        Exit Function
    End If
    
    'Disable screen updating, events and recalc
    Call Echo(False)
        Call Clear
        If Not NewRecord Then
            Call DeleteRecordData(strRecord)
            DeleteRecord = True
        End If
        
        'Update record properties
        mbolNewRecord = True
        mbolDirty = False
        
        'Raise events AfterDelete, Current and BeforeSaveWorkbook
        RaiseEvent AfterDelete(strRecord)
        RaiseEvent Current
        'Save workbook after deletion?
        If Me.AutoSaveWorkbook Then
            RaiseEvent BeforeSaveWorkbook(intCancelSaveWorkbook)
            If Not intCancelSaveWorkbook Then
                ThisWorkbook.Save
            End If
        End If
    'Enabled screen updating, events and recalc1
    Call Echo(True)
End Function

Private Sub DeleteRecordData(strRecord As String)
    Dim rg As Range
    Dim rgRecord As Range
    Dim strRecordRange As String
    Dim lngLastRow As Long
    Dim lngSafeRow As Long
    Dim intColumns As Integer
    
    'Get the last row used by Database parameters
    lngSafeRow = mintdbRecordsFirstRow + mconDataBaseParameters
    mws.Unprotect
        'Delete record entry
        Set rg = mws.Range(mstrdbSavedRecords)
        lngLastRow = rg.Row + rg.Rows.Count - 1
        'Set the last safe sheet row to delete entire row
        If lngSafeRow < lngLastRow Then
            lngSafeRow = lngLastRow
        End If
        
        'Delete the One-side record from mconSavedRecords range
        Set rgRecord = rg.Find(strRecord)
        intColumns = mintdbRangeOffset + mintdbOneSideColumnsCount
        rgRecord.Resize(1, intColumns).ClearContents
        
        If rgRecord.Row <> lngLastRow Then
            'Reposition other record entries by copy and paste
            mws.Range(Cells(rgRecord.Row + 1, rgRecord.Column), Cells(lngLastRow, rgRecord.Column + intColumns - 1)).Copy
            rgRecord.PasteSpecial xlPasteValues
        End If
        
        'Clear last mstrdbSavedRecords record row
        mws.Range(Cells(lngLastRow, rgRecord.Column), Cells(lngLastRow, rgRecord.Column + intColumns - 1)).ClearContents
        'Resize mstrdbSavedRecords range name without deleted Record
        rg.Resize(rg.Rows.Count - 1).Name = "'" & mws.Name & "'!" & mstrdbSavedRecords
        
        'Delete de Many-side records and it range name
        strRecordRange = mstrdbManySidePrefix & FixName(strRecord)
        Set rg = mws.Range(strRecordRange)
        'Verify if record data, mstrdbSavedRecords or the Database parameters range use the same rows
        If rg.Row <= lngSafeRow Then
            'This saved records data rows must just be cleaned
            rg.Resize(mintdbManySideRowsCount, mintdbManySideColumnsCount).ClearContents
        Else
            'It is safe to delete entire saved records data rows
            rg.Resize(mintdbManySideRowsCount).EntireRow.Delete
            'Provision to keep rows hidden
            mws.Range(Cells(mintdbRecordsFirstRow, 1), Cells(mws.Rows.Count, 1)).EntireRow.Hidden = True
        End If
        'Delete the many-records Range name
        mws.Names(strRecordRange).Delete
        'Scroll to row 1
        ActiveWindow.ScrollRow = 1
    If mbolSheetProtected Then
        mws.Protect
    End If
End Sub

Private Function ManySideRecordsSize(Optional Rows As Integer, Optional Cols As Integer) As Variant
    Dim rgRelation As Range
    Dim rgArea As Range
    Dim strRelation As String
    Dim intI As Integer
    Dim intRows As Integer
    Dim intMaxRows As Integer
    Dim intMaxCols As Integer
    Dim intCols As Integer
    Dim intCol As Integer
    
    'Define the many-side records dimension with intRows, intCols
    For intI = 1 To 4
        strRelation = Choose(intI, mstrdbManySide1, mstrdbManySide2, mstrdbManySide3, mstrdbManySide4)
        If Len(strRelation) Then
            Set rgRelation = mws.Range(strRelation)
            For Each rgArea In rgRelation.Areas
                For intCol = 1 To rgArea.Columns.Count
                    If rgArea.Cells(1, intCol).MergeCells Then
                        intCol = intCol + rgArea.Cells(1, intCol).MergeArea.Columns.Count - 1
                    End If
                    intMaxCols = intMaxCols + 1
                Next
                
                If rgArea.Rows.Count > intMaxRows Then
                    intMaxRows = rgArea.Rows.Count
                End If
            Next
            intRows = intRows + intMaxRows + 1
            If intMaxCols > intCols Then
                intCols = intMaxCols
            End If
            intMaxRows = 0
            intMaxCols = 0
        End If
    Next
    
    Rows = intRows
    Cols = intCols
    ManySideRecordsSize = Array(intRows, intCols)
End Function

Private Function FixName(ByVal strName As String, Optional LastErrorNumber As Long)
    'Invalid characters inside range names: @#$%&()+~`"':;,.|!?_-/\*[]{}
    Dim strInvalidChars As String
    Dim strChar As String
    Dim intI As Integer
    Const errNameInvalid = 1004
    
    'Search for invalid characters
    strInvalidChars = "@#$%&()+~`´':;,.|!?-/\*[]{}" & """"
    For intI = 1 To Len(strInvalidChars)
        'Get each invalid character and take it out
        strChar = Mid(strInvalidChars, intI, 1)
        strName = Replace(strName, strChar, "")
    Next
    
    'Now change spaces to underscores
    strName = Replace(strName, " ", "_")
    
    FixName = strName
End Function

Public Sub ShowRecord(Record As RecordPosition)
    Dim rg As Range
    Dim strRecord As String
    Dim lngFirstRec As Long
    Dim lngLastRec As Long
    Dim bolMoveRecord As Boolean
    
    lngFirstRec = mws.Range(mstrdbSavedRecords).Row + 1
    lngLastRec = mws.Range(mstrdbSavedRecords).Row + mws.Range(mstrdbSavedRecords).Rows.Count - 1
        
    If mbolNewRecord And Record = NextRec Then Record = FirstRec
    If mbolNewRecord And Record = PreviousRec Then Record = LastRec
    
    Select Case Record
        Case FirstRec, LastRec, NewRec
            Set rg = mws.Range(mstrdbSavedRecords)
            Select Case Record
                Case FirstRec
                    Set rg = rg.Cells(2)
                Case LastRec
                    Set rg = rg.Cells(rg.Rows.Count)
                    If rg = "" Then
                        Do
                            Set rg = rg.Offset(-1)
                        Loop Until Len(rg) > 0
                    End If
                Case NewRec
                    Set rg = rg.Cells(1)
            End Select
            bolMoveRecord = True
        Case PreviousRec, NextRec
            strRecord = mws.Range(mstrdbDataValidationList)
            Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord, , , xlWhole)
            If Record = NextRec And rg.Row < lngLastRec Then
                Set rg = rg.Offset(1)
                bolMoveRecord = True
            ElseIf Record = PreviousRec And rg.Row > lngFirstRec Then
                Set rg = rg.Offset(-1)
                bolMoveRecord = True
            End If
    End Select
    
    If bolMoveRecord Then
        'Move to selected record!
        mws.Range(mstrdbDataValidationList) = rg.Value
    End If
End Sub

Public Function CopyRecord(strRecord As String, rgOneSide As Range, rgManySide As Range) As Boolean
    Dim rg As Range
    
    Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord, , , xlWhole)
    If Not rg Is Nothing Then
        If Len(mstrdbOneSide) Then
            Set rgOneSide = Range(rg.Offset(0, mintdbRangeOffset), rg.Offset(0, mintdbRangeOffset + mintdbOneSideColumnsCount - 1))
        End If
        
        If Len(mstrdbManySide1) Then
            Set rg = mws.Range(mstrdbManySidePrefix & FixName(strRecord))
            Set rgManySide = Range(rg.Offset(0, 0), rg.Offset(mintdbManySideRowsCount - 2, mintdbManySideColumnsCount - 1))
        End If
        CopyRecord = True
    End If
End Function

Public Function PasteRecord(strRecord As String, _
                            rgOneSide As Range, _
                            rgManySide As Range, _
                            Optional PasteAsNewRecord As Boolean) As Boolean
    Dim rg As Range
    Dim strRangeName As String
    Dim lngRow As Long
    Dim intI As Integer
    Dim bolProtect As Boolean
    Dim bolWorksheetIsFull As Boolean
    Dim bolRecordPaste As Boolean
    
    Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord, , , xlWhole)
    If Not rg Is Nothing And PasteAsNewRecord Then
        'Add a name count suffix to paste existing record as new one
        Do
            'Find a new record name
            intI = intI + 1
            Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord & intI, , , xlWhole)
        Loop Until rg Is Nothing
        strRecord = strRecord & intI
    End If
    
    Call Echo(False)
    bolProtect = mws.ProtectContents
    mws.Unprotect
    strRangeName = mstrdbManySidePrefix & FixName(strRecord)
    If rg Is Nothing Then
        'strRecord does not exist. Createt it!
        'Define sheet row where next Record data will be stored
        lngRow = NextEntryRow(bolWorksheetIsFull)

        'Verify if sheet is full
        If bolWorksheetIsFull Then
            'No more room to save data
            MsgBox "There is no more room to paste records", vbCritical, "Workdhseet database is full"
            Exit Function
        End If
        
        'Verify if mstrSavedRecords last rows is a empty cell
        Set rg = mws.Range(mstrdbSavedRecords)
        If Not rg.Cells(rg.Rows.Count, 1) = "" Then
            'Insert a new row at bottom of SavedRecords range name and update rg object
            rg.Resize(rg.Rows.Count + 1).Name = "'" & mws.Name & "'!" & mstrdbSavedRecords
            Set rg = mws.Range(mstrdbSavedRecords)
        End If
        
        'Position on new cell of SavedRecords range and save New Record name
        Set rg = rg.Cells(rg.Rows.Count, 1)
        rg = strRecord
        
        If Len(mstrdbManySide1) Then
            'Define Record name as 'rec_<strRecord>' and create it range name
            mws.Names.Add strRangeName, "='" & mws.Name & "'!" & mstrdbManySideFirstColumn & lngRow
            mws.Names(strRangeName).Visible = False
        End If
    End If

    If Len(mstrdbOneSide) Then
        'Paste the one side record
        Set rg = rg.Offset(0, mintdbRangeOffset)
        rgOneSide.Copy
        rg.PasteSpecial xlPasteValues
        bolRecordPaste = True
    End If
    
    If Len(mstrdbManySide1) Then
        'Paste the Many side records
        Set rg = mws.Range(strRangeName)
        rgManySide.Copy
        rg.PasteSpecial xlPasteValues
        bolRecordPaste = True
    End If
    
    If bolProtect Then mws.Protect
    Call Echo(True)
    PasteRecord = bolRecordPaste
End Function

Public Sub Sort()
    Dim rg As Range
    Dim bolProtect As Boolean
    
    Call Echo(False)
    bolProtect = mws.ProtectContents
    mws.Unprotect
        Set rg = mws.Range(mstrdbSavedRecords)
        'Sort SavedRecords and find strRecord position
        Set rg = mws.Range(mws.Cells(rg.Row + 1, rg.Column), _
                           mws.Cells(rg.Row + rg.Rows.Count - 1, rg.Column + mintdbRangeOffset + mintdbOneSideColumnsCount - 1))
        'Unhide range rows because Sort does not works well on hidden rows
        rg.EntireRow.Hidden = False
            rg.Sort rg.Cells(, 1)
        rg.EntireRow.Hidden = True
    If bolProtect Then mws.Protect
    Call Echo(True)
End Sub


