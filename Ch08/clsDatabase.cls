VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mwb As Workbook
Attribute mwb.VB_VarHelpID = -1
Private WithEvents mws As Worksheet
Attribute mws.VB_VarHelpID = -1
Public Dirty As Boolean  'Indicate if record dat had been changed
Private mstrLastRecord As String 'Retain the name of current record
Private Enum Operation
    LoadRecord = 1
    SaveRecord = 2
End Enum

Dim mstrdbDataValidationList As String
Dim mstrdbSavedRecords As String
Dim mstrdbRecordName As String
Dim mstrdbOneSide As String
Dim mintdbOneSideColumsCount As Integer
Dim mstrdbManySide1 As String
Dim mstrdbManySide2 As String
Dim mstrdbManySide3 As String
Dim mstrdbManySide4 As String
Dim mstrdbManySidePrefix As String
Dim mintdbManySideColumnsCount As Integer
Dim mintdbManySideRowsCount As Integer
Dim mintdbRecordsFirstRow As Integer
Dim mstrdbManySideFirstColumn As String
Dim mintdbRangeOffset As Integer

Private Sub Class_Initialize()
    Set mwb = ThisWorkbook
    Set mws = ActiveSheet
    mstrdbDataValidationList = mws.Range("dbDataValidationList")
    mstrdbSavedRecords = mws.Range("dbSavedRecords")
    mstrdbRecordName = mws.Range("dbRecordName")
    mstrdbOneSide = mws.Range("dbOneSide")
    mintdbOneSideColumsCount = mws.Range("dbOneSideColumsCount")
    mstrdbManySide1 = mws.Range("dbManySide1")
    mstrdbManySide2 = mws.Range("dbManySide2")
    mstrdbManySide3 = mws.Range("dbManySide3")
    mstrdbManySide4 = mws.Range("dbManySide4")
    mstrdbManySidePrefix = mws.Range("dbManySidePrefix")
    mintdbManySideColumnsCount = mws.Range("dbManySideColumnsCount")
    mintdbManySideRowsCount = mws.Range("dbManySideRowsCount ")
    mintdbRecordsFirstRow = mws.Range("dbRecordsFirstRow")
    mstrdbManySideFirstColumn = mws.Range("dbManySideFirstColumn")
    mintdbRangeOffset = mws.Range("dbRangeOffset")
End Sub

Private Sub mwb_BeforeClose(Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    Dim strRecord As String
    Dim bolSaved As Boolean
    
    If Me.Dirty Then
        strRecord = mws.Range(mstrdbDataValidationList)
        If strRecord = "" Then strRecord = "New " & mstrdbRecordName
        strTitle = "Save " & strRecord & " data?"
        strMsg = strRecord & " data had been changed." & vbCrLf
        strMsg = strMsg & "Save " & strRecord & " data before close the workbook?"
        Select Case MsgBox(strMsg, vbYesNoCancel + vbQuestion, strTitle)
            Case vbYes
                bolSaved = Save(strRecord)
                Cancel = Not bolSaved
            Case vbCancel
                Cancel = True
        End Select
    End If
End Sub

Private Sub mws_SelectionChange(ByVal Target As Range)
    If mws.Range(mstrdbDataValidationList) = "" Then
        mstrLastRecord = "New " & mstrdbRecordName
    Else
        mstrLastRecord = mws.Range(mstrdbDataValidationList)
    End If
End Sub

Private Sub mws_Change(ByVal Target As Range)
    Select Case Target.Address
        Case Is = mws.Range(mstrdbDataValidationList).Address
            'User is trying to load a new Record
            TryToLoadSelectedRecord
        Case Else
            'Sheet data has changed
            Me.Dirty = True
            If mws.Range(mstrdbDataValidationList) = "New " & mstrdbRecordName Then
                Application.EnableEvents = False
                    mws.Range(mstrdbDataValidationList) = ""
                Application.EnableEvents = True
            End If
     End Select
End Sub

Private Sub TryToLoadSelectedRecord()
    Dim ws As Worksheet
    Dim strMsg As String
    Dim strNewRecord As String
    
    Set ws = ActiveSheet
    strNewRecord = ws.Range(mstrdbDataValidationList)
    
    'Verify if current Record had been changed
    If Me.Dirty Then
        'Save current Record before change it?
        strMsg = mstrLastRecord & " data had been changed." & vbCrLf & vbCrLf
        strMsg = strMsg & "Save " & mstrLastRecord & " before load '" & strNewRecord & "'?"
        If MsgBox(strMsg, vbQuestion + vbYesNo, "Save current data?") = vbYes Then
            If Not Save(mstrLastRecord) Then
                'Record data not saved!
                Application.EnableEvents = False
                    ws.Range(mstrdbDataValidationList) = mstrLastRecord
                Application.EnableEvents = True
                Exit Sub
            End If
        End If
        Me.Dirty = False
    End If
        
    'Load selected Record data
    Call Load(strNewRecord)
End Sub

Private Sub Load(strRecord As String)
    'Disable screen updating, events and recalc
    SetScreenEventsRecalc (False)
        Select Case strRecord
            Case "", "New " & mstrdbRecordName
                'User selected a "New Record"
                Call Clear
                mws.Range(mstrdbDataValidationList) = "New " & mstrdbRecordName
            Case Else
                Call LoadSaveData(strRecord, LoadRecord)
                mws.Range(mstrdbDataValidationList).Select
                mstrLastRecord = strRecord
        End Select
        Me.Dirty = False
    'Enable screen updating, events and recalc
    SetScreenEventsRecalc (True)
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

Public Function Save(Optional strLastRecord As String) As Boolean
    Dim rg As Range
    Dim strRecord As String
    Dim bolNewRecord As Boolean
    Dim bolRecordSaved As Boolean
    
    'Verify if Record data is still empty
    strRecord = mws.Range(mstrdbDataValidationList)
    If strRecord = "New " & mstrdbRecordName Then
        Exit Function
    End If

    If strLastRecord = "" Then
        strLastRecord = strRecord
    End If
    strRecord = GetRecordName(strLastRecord, bolNewRecord)
    
    If Len(strRecord) Then
        'Disable application events to allow cell change by macro code
        SetScreenEventsRecalc (False)
            mws.Unprotect
                bolRecordSaved = SaveData(strRecord, bolNewRecord)
            mws.Protect
            If bolRecordSaved Then
                'Define current Record as saved Record
                mws.Range(mstrdbDataValidationList) = strRecord
                mws.Range(mstrdbDataValidationList).Select
                
                'Save the worbook
                ThisWorkbook.Save
                mstrLastRecord = strRecord
                Me.Dirty = False
                Save = True
                MsgBox mstrdbRecordName & " data had been saved as '" & strRecord & "'!", , "BMI Companion Chart"
            Else
                MsgBox "There is no more room to save data on this worksheet!", vbCritical, "Can't save data"
            End If
        SetScreenEventsRecalc (True)
   End If
End Function

Private Function SaveData(strRecord As String, bolNewRecord As Boolean) As Boolean
    Dim rg As Range
    Dim strRangeName As String
    Dim strAddress As String
    Dim lngRow As Long
    Dim bolWorksheetIsFull As Boolean
    
    Set rg = mws.Range(mstrdbSavedRecords)
    If bolNewRecord Then
        'Define sheet row where next Record data will be stored
        lngRow = NextEntryRow(bolWorksheetIsFull)
        
        'Verify if sheet can receive more records
        If bolWorksheetIsFull Then
            'No more room to save data
            Exit Function
        End If
        
        'Insert a new row at bottom of mstrdbSavedRecords range name and update rg object
        rg.Resize(rg.Rows.Count + 1).Name = "'" & mws.Name & "'!" & mstrdbSavedRecords
        Set rg = mws.Range(mstrdbSavedRecords)

        'Position on new cell of mstrdbSavedRecords range and save new Record name
        rg.Cells(rg.Rows.Count, 1) = strRecord
        
        If Len(mstrdbManySide1) Then
            'Define many-side Record name as 'mstrdbManySidePrefix<strRecord>' and create tbe range name
            strRangeName = mstrdbManySidePrefix & FixName(strRecord)
            strAddress = "='" & mws.Name & "'!" & mstrdbManySideFirstColumn & lngRow
            mws.Names.Add strRangeName, strAddress, False
        End If
    End If
        
    Call LoadSaveData(strRecord, SaveRecord)
    
    'Sort mstrdbSavedRecords range keeping "New <mstrdbRecordName>"  on the top of the list
    Set rg = mws.Range(Cells(rg.Row + 1, rg.Column), _
                      Cells(rg.Row + rg.Rows.Count, rg.Column + mintdbRangeOffset + mintdbOneSideColumsCount))
    rg.EntireRow.Hidden = False
        rg.Sort rg.Cells(, 1)
    rg.EntireRow.Hidden = True

    mws.Range("A1").Select
    SaveData = True
End Function

Private Function GetRecordName(strRecord As String, bolNewRecord As Boolean) As String
    Dim rg As Range
    Dim strNewRecord As String
    Static sintDefaultName As Integer
    
    If strRecord = "" Then
        sintDefaultName = sintDefaultName + 1
        strRecord = "New " & mstrdbRecordName & " " & Replace(Date, "/", "_")
        If sintDefaultName > 1 Then
            strRecord = strRecord & " " & sintDefaultName
        End If
        strRecord = InputBox("Data will be saved as:", "Confirm data name", strRecord)
    End If
    
    If Len(strRecord) Then
        'Verify if strRecord already exist on mstrdbSavedRecords
        Set rg = mws.Range(mstrdbSavedRecords).Find(strRecord)
    
        If rg Is Nothing Then
            bolNewRecord = True
        Else
            'Confirm proposed record name
            strNewRecord = InputBox(mstrdbRecordName & " '" & strRecord & "' already exist. Do you want to overwrite it?", _
                                    "Overwrite " & strRecord & " data?", strRecord)
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
            lngRow = mintdbRecordsFirstRow + (mws.Range(mstrdbSavedRecords).Rows.Count * mintdbManySideRowsCount)
        End If
        bolWorksheetIsFull = (lngRow > (mws.Rows.Count - mintdbManySideRowsCount))
    Else
        'Just one-side record to find next entry row
        lngRow = mintdbRecordsFirstRow + mws.Range(mstrdbSavedRecords).Rows.Count
        bolWorksheetIsFull = (lngRow > (mws.Rows.Count - mws.Range(mstrdbSavedRecords).Rows.Count))
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

Public Function DeleteRecord() As Boolean
    Dim strRecord As String
    Dim strMsg As String
    Dim strTitle As String
    Dim intCancelDelete As Integer
    Dim intCancelSave As Integer
    Dim bolNewRecord As Boolean
    
    strRecord = mws.Range(mstrdbDataValidationList)
    If strRecord = "" Or strRecord = "New " & mstrdbRecordName Then
        If Me.Dirty Then
            bolNewRecord = True
            strMsg = "New " & mstrdbRecordName & " data has not been saved yet." & vbCrLf
            strMsg = strMsg & "Do you want to delete it?"
            strTitle = "Delete unsaved record?"
        Else
            Exit Function
        End If
    Else
        strMsg = "Do you want to delete " & strRecord & " record?"
        strTitle = "Delete record?"
    End If
        
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion, strTitle) = vbYes Then
        'Disable screen updating, events and recalc
        SetScreenEventsRecalc (False)
            Call Clear
            If Not bolNewRecord Then
                Call DeleteRecordData(strRecord)
            End If
            DeleteRecord = True
            Me.Dirty = False
            mstrLastRecord = "New " & mstrdbRecordName
        'Enabled screen updating, events and recalc
        SetScreenEventsRecalc (True)
        
        mws.Range(mstrdbDataValidationList) = mstrLastRecord
        
        'Save workbook after deletion
        ThisWorkbook.Save
    End If
End Function

Private Sub DeleteRecordData(strRecord As String)
    Dim rg As Range
    Dim rgRecord As Range
    Dim strRecordRange As String
    Dim lngLastRow As Long
    Dim lngSafeRow As Long
    Dim intColumns As Integer
    
    mws.Unprotect
        Set rg = mws.Range(mstrdbSavedRecords)
        'Get the last row used by Database parameters
        lngSafeRow = mintdbRecordsFirstRow
        lngLastRow = rg.Row + rg.Rows.Count - 1
        'Set the last safe sheet row to delete entire row
        If lngSafeRow < lngLastRow Then
            lngSafeRow = lngLastRow
        End If
        
        'Delete the One-side record from mstrdbSavedRecords range
        Set rgRecord = rg.Find(strRecord)
        intColumns = mintdbRangeOffset + mintdbOneSideColumsCount
        rgRecord.Resize(1, intColumns).ClearContents
        
        If rgRecord.Row <> lngLastRow Then
            'Reposition other record entries by copy and paste
            mws.Range(Cells(rgRecord.Row + 1, rgRecord.Column), Cells(lngLastRow, rgRecord.Column + intColumns - 1)).Copy
            rgRecord.PasteSpecial xlPasteValues
        End If
        
        'Clear last mstrdbSavedRecords record row
        Range(Cells(lngLastRow, rgRecord.Column), Cells(lngLastRow, rgRecord.Column + intColumns - 1)).ClearContents
        'Resize mstrdbSavedRecords range name without deleted Record
        rg.Resize(rg.Rows.Count - 1).Name = "'" & mws.Name & "'!" & mstrdbSavedRecords
        
        'Delete the Many-side records and it range name
        strRecordRange = mstrdbManySidePrefix & FixName(strRecord)
        Set rg = mws.Range(strRecordRange)
        'Verify if record data amd mstrdbSavedRecords range use the same rows
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
    mws.Protect
End Sub






