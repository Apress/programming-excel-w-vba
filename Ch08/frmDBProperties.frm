VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDBProperties 
   Caption         =   "Database Properties"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   OleObjectBlob   =   "frmDBProperties.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDBProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Enum ShowTab
    PreviousTab = -1
    NextTab = 1
End Enum

Private Enum DataBaseOperation
    CreateDatabase = 1
    RemoveDatabase = 2
End Enum

Private Sub chkCreateNavigationButtons_Click()
    Me.cmdRecordPosition.Enabled = Me.chkCreateNavigationButtons
    Me.txtRecordPosition.Enabled = Me.chkCreateNavigationButtons
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim rg As Range
    Dim strName As String
    Dim strNameScope As String
    Dim intI As Integer
    Const conNormalWidth = 473
    Const conWhite = &HFFFFFF
    
    On Error Resume Next
    
    Me.Width = conNormalWidth
    Application.EnableEvents = False
    Set ws = Application.ActiveSheet
    strNameScope = "'" & ws.Name & "'!"
    Me.tabControl.Style = fmTabStyleNone
    
    Set rg = Range(strNameScope & "dbDataValidationList")
    If rg Is Nothing Then
        Me.tabControl.Value = 0
        Me.cmdPrevious.Visible = True
        Me.cmdNext.Visible = True
        Me.txtdbRecordsFirstRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count + 3
        Me.txtRecordPosition = SetRecordPosition()
        Call LoadNames
    Else
        Me.cmdDefine.Caption = "Close"
        Me.cmdDefine.Accelerator = "C"
        Me.cmdDefine.Enabled = True
        Me.cmdCancel.Caption = "Remove"
        Me.cmdCancel.Accelerator = "R"
        Me.cmdPrevious.Visible = False
        Me.cmdNext.Visible = False
        Me.txtdbRecordName1.Locked = True
        Me.txtdbRecordName1.BackColor = conWhite
        Me.txtdbManySidePrefix.Locked = True
        Me.txtdbManySidePrefix.BackColor = conWhite
        
        'Update UserForm TextBoxes
        For intI = 1 To 15
            strName = Choose(intI, "dbRecordName", _
                                   "dbDataValidationList", _
                                   "dbSavedRecords", _
                                   "dbRecordsFirstRow", _
                                   "dbOneSide", _
                                   "dbOneSideColumnsCount", _
                                   "dbManySide1", _
                                   "dbManySide2", _
                                   "dbManySide3", _
                                   "dbManySide4", _
                                   "dbManySideFirstColumn", _
                                   "dbManySideColumnsCount", _
                                   "dbManySideRowsCount", _
                                   "dbManySidePrefix", _
                                   "dbRangeOffset")
            Me("txt" & strName) = ws.Range(strNameScope & strName)
        Next
        
        Me.tabControl.Value = Me.tabControl.Pages.Count - 1
        Call CalculateManySideRecords
    End If
End Sub

Private Sub UserForm_Terminate()
    Application.EnableEvents = True
End Sub

Private Sub LoadNames(Optional bolAllRangeNames As Boolean)
    Dim obj As Object
    Dim nm As Name
    Dim varNames() As Variant
    Dim intI As Integer
    
    On Error Resume Next
    
    'Load desired names on varNames() array
    If bolAllRangeNames Then
        Set obj = ThisWorkbook
    Else
        Set obj = ActiveSheet
    End If
    
    ReDim varNames(obj.Names.Count - 1)
    
    For Each nm In obj.Names
        varNames(intI) = nm.Name
        intI = intI + 1
    Next
    
    'Populate Comboboxes
    Me.cbodbOneSide.List = varNames()
    Me.cbodbManySide1.List = varNames()
    Me.cbodbManySide2.List = varNames()
    Me.cbodbManySide3.List = varNames()
    Me.cbodbManySide4.List = varNames()
End Sub

Private Sub chkWorkbookNames_Click()
    Call LoadNames(Me.chkWorkbookNames)
End Sub

Private Sub cmdNext_Click()
    Call ShowPage(NextTab)
End Sub

Private Sub cmdPrevious_Click()
    Call ShowPage(PreviousTab)
End Sub

Private Sub ShowPage(Action As ShowTab)
    Static sintPage As Integer
    Dim intMaxPages As Integer

    If Action = NextTab Then
        If Not ValidatePage(sintPage) Then Exit Sub
    End If
    
    sintPage = sintPage + Action
    intMaxPages = Me.tabControl.Pages.Count - 1

    If sintPage < 0 Then sintPage = 0
    If sintPage > intMaxPages Then sintPage = intMaxPages
    Me.tabControl.Value = sintPage
    Me.cmdDefine.Enabled = (sintPage = Me.tabControl.Pages.Count - 1)
    Me.chkWorkbookNames.Visible = (sintPage = 2 Or sintPage = 3)
End Sub

Function ValidatePage(intPage As Integer) As Boolean
    Dim strMsg As String
    Dim strTitle As String
    Dim bolValidateFail As Boolean
    
    Select Case intPage
        Case 0
            'Validata record name
            If Len(Me.txtdbRecordName) = 0 Then
                strMsg = "Define the default name for the worksheet record."
                strTitle = "Record name?"
                bolValidateFail = True
            End If
        Case 1
            'Validata Data Validation list
            If Len(Me.txtdbDataValidationList) = 0 Then
                strMsg = "Select a cell for the Records Data Validation list and try again."
                strTitle = "Data Validation list cell?"
                bolValidateFail = True
            End If
        Case 3
            'Validata OneSide and ManySide records
            If Me.txtdbOneSideColumnsCount = 0 And Me.txtdbManySideRowsCount = 0 Then
                strMsg = "Select the One-Side and/or the Many-Side cells that define the worksheet records ranges!"
                strTitle = "Select cells to be saved as worksheet records"
                bolValidateFail = True
            End If
    End Select
    
    If bolValidateFail Then
        MsgBox strMsg, vbQuestion, strTitle
    Else
        ValidatePage = True
    End If
End Function

Private Sub tabControl_Change()
    Dim lngRec As Long
    
    On Error Resume Next
    
    If Me.txtdbManySideRowsCount = 0 Then
        lngRec = (ActiveSheet.Rows.Count - Me.txtdbRecordsFirstRow)
    Else
        lngRec = (ActiveSheet.Rows.Count - Me.txtdbRecordsFirstRow) / Me.txtdbManySideRowsCount
    End If
    Me.lblNumRecords.Caption = lngRec & " records allowed"
    Me.cmdPrevious.Enabled = (Me.tabControl.Value > 0)
    Me.cmdNext.Enabled = (Me.tabControl.Value < Me.tabControl.Pages.Count - 1)
End Sub

Private Sub txtdbRecordName_Change()
    Me.txtdbRecordName1 = Me.txtdbRecordName
End Sub

Private Sub cmdDataValidationList_Click()
    Dim varFormula As Variant
    Dim varName As Variant
    Dim strListRange As String
    Dim strRange As String
    
    On Error Resume Next
    
    Me.Hide
    strRange = GetRange("Select cell for the Data Validation list:", "Data Validation List?", Me.txtdbDataValidationList)
    If Len(strRange) Then
        varName = Range(strRange).Name.Name
        If Len(varName) Then
            strRange = varName
        End If
        Me.txtdbDataValidationList = strRange
        Range(strRange).Merge
        
        'Verify if selected range has a data validation list
        strListRange = Range(strRange).Validation.Formula1
        If Len(strListRange) Then
            Me.txtdbSavedRecords = Mid(strListRange, 2)
        Else
            Me.txtdbSavedRecords = "SavedRecords"
        End If
    End If
    Me.Show
End Sub

Private Function GetRange(strMsg As String, strTitle As String, Optional Default As Variant) As String
    Dim varRg As Variant
    Dim rgArea As Range
    Dim strAddress As String
    Dim bolInvalidSelection As Boolean
    Const conRange = 8
        
    On Error Resume Next
    
    Set varRg = Application.InputBox(strMsg, strTitle, Default, , , , , conRange)
    If IsObject(varRg) Then
        For Each rgArea In varRg.Areas
            If Len(strAddress) Then
                strAddress = strAddress & ","
            End If
            strAddress = strAddress & rgArea.Address
        Next
        GetRange = strAddress
    End If
End Function

Private Sub txtdbDataValidationList_Change()
    Me.txtdbDataValidationList1 = Me.txtdbDataValidationList
End Sub

Private Sub cmddbOneSide_Click()
    Dim strRange As String
    Dim strMsg As String
        
    On Error Resume Next
    
    Me.Hide
    strMsg = "Select all cells that belongs to the 'one side' of the worksheet record." & vbCrLf
    strRange = GetRange(strMsg, "One-side record sheet cells", Me.cbodbOneSide)
    If Len(strRange) Then
       Me.cbodbOneSide = strRange
    End If
    Me.Show
End Sub

Private Sub cbodbOneSide_Change()
    Dim intCells As Integer
    Dim strAddress As String
    Const conColsDatabase = 6
    
    If IsRange(Me.cbodbOneSide) Then
        'Count cells selected
        intCells = CalculateOneSideColumns()
        Me.txtdbOneSideColumnsCount = intCells
        'Define save column for many-side records
        strAddress = Cells(1, intCells + conColsDatabase).Address
        strAddress = Left(strAddress, InStrRev(strAddress, "$"))
        Me.txtdbManySideFirstColumn = strAddress
    Else
        Me.cbodbOneSide = ""
        Me.txtdbOneSideColumnsCount = 0
        Me.txtdbManySideFirstColumn = ""
    End If
    
    If Left(Me.cbodbOneSide, 1) = "'" Then
        Me.txtdbOneSide = "'" & Me.cbodbOneSide
    Else
        Me.txtdbOneSide = Me.cbodbOneSide
    End If
    Me.cmdClearcbodbOneSide.Enabled = (Len(Me.cbodbOneSide) > 0)
End Sub

Private Function IsRange(strRange As String) As Boolean
    Dim rg As Range
    
    On Error Resume Next
    Set rg = Range(strRange)
    IsRange = (Err = 0)
End Function

Private Function CalculateOneSideColumns() As Integer
    Dim rg As Range
    Dim rgArea As Range
    Dim strAddress As String
    Dim intNumCols As Integer
    Dim intI As Integer
    Dim intJ As Integer
    
    Set rg = Range(Me.cbodbOneSide)
    For Each rgArea In rg.Areas
        For intI = 1 To rgArea.Rows.Count
            For intJ = 1 To rgArea.Columns.Count
                If rgArea.Cells(intI, intJ).MergeCells Then
                    intI = intI + rgArea.Cells(intI, intJ).MergeArea.Rows.Count - 1
                    intJ = intJ + rgArea.Cells(intI, intJ).MergeArea.Columns.Count - 1
                End If
                intNumCols = intNumCols + 1
            Next intJ
        Next intI
    Next
    CalculateOneSideColumns = intNumCols
End Function

Private Sub txtdbOneSideColumnsCount_Change()
    Me.txtdbOneSideColumnsCount1 = Me.txtdbOneSideColumnsCount
End Sub

Private Sub cmdClearcbodbOneSide_Click()
    Me.cbodbOneSide = ""
End Sub

Private Sub cmddbManySide1_Click()
    Call GetdbManySide(1)
End Sub

Private Sub cmddbManySide2_Click()
    Call GetdbManySide(2)
End Sub

Private Sub cmddbManySide3_Click()
    Call GetdbManySide(3)
End Sub

Private Sub cmddbManySide4_Click()
    Call GetdbManySide(4)
End Sub

Private Sub cbodbManySide1_Change()
    If IsRange(Me.cbodbManySide1) Then
        If Left(Me.cbodbManySide1, 1) = "'" Then
            Me.txtdbManySide1 = "'" & Me.cbodbManySide1
        Else
            Me.txtdbManySide1 = Me.cbodbManySide1
        End If
    Else
        Me.cbodbManySide1 = ""
        Me.txtdbManySide1 = ""
    End If
    
    Me.cmdClear1.Enabled = (Len(Me.cbodbManySide1) > 0)
    Me.cbodbManySide2.Enabled = (Len(Me.cbodbManySide1) > 0)
    Me.cmddbManySide2.Enabled = (Len(Me.cbodbManySide1) > 0)
    Call CalculateManySideRecords
End Sub

Private Sub cbodbManySide2_Change()
    If IsRange(Me.cbodbManySide2) Then
        If Left(Me.cbodbManySide2, 1) = "'" Then
            Me.txtdbManySide2 = "'" & Me.cbodbManySide2
        Else
            Me.txtdbManySide2 = Me.cbodbManySide2
        End If
    Else
        Me.cbodbManySide2 = ""
        Me.txtdbManySide2 = ""
    End If
    
    Me.cmdClear2.Enabled = (Len(Me.cbodbManySide2) > 0)
    Me.cbodbManySide3.Enabled = (Len(Me.cbodbManySide2) > 0)
    Me.cmddbManySide3.Enabled = (Len(Me.cbodbManySide1) > 0 And Len(Me.cbodbManySide2) > 0)
    Call CalculateManySideRecords
End Sub

Private Sub cbodbManySide3_Change()
    If IsRange(Me.cbodbManySide3) Then
        If Left(Me.cbodbManySide3, 1) = "'" Then
            Me.txtdbManySide3 = "'" & Me.cbodbManySide3
        Else
            Me.txtdbManySide3 = Me.cbodbManySide3
        End If
    Else
        Me.cbodbManySide3 = ""
        Me.txtdbManySide3 = ""
    End If
    
    Me.cmdClear3.Enabled = (Len(Me.cbodbManySide3) > 0)
    Me.cbodbManySide4.Enabled = (Len(Me.cbodbManySide3) > 0)
    Me.cmddbManySide4.Enabled = (Len(Me.cbodbManySide1) > 0 And Len(Me.cbodbManySide2) > 0 And Len(Me.cbodbManySide3) > 0)
    Call CalculateManySideRecords
End Sub

Private Sub cbodbManySide4_Change()
    If IsRange(Me.cbodbManySide4) Then
        If Left(Me.cbodbManySide4, 1) = "'" Then
            Me.txtdbManySide4 = "'" & Me.cbodbManySide4
        Else
            Me.txtdbManySide4 = Me.cbodbManySide4
        End If
    Else
        Me.cbodbManySide4 = ""
        Me.txtdbManySide4 = ""
    End If
    
    Me.cmdClear4.Enabled = (Len(Me.cbodbManySide4) > 0)
    Call CalculateManySideRecords
End Sub

Private Function CalculateManySideRecords() As Integer
    Dim rg As Range
    Dim rgArea As Range
    Dim strCtl As String
    Dim strRange As String
    Dim intI As Integer
    Dim intJ As Integer
    Dim intMaxRows As Integer
    Dim intNumRows As Integer
    Dim intMaxCols As Integer
    Dim intNumCols As Integer
    Dim intPos As Integer
    Dim intPos2 As Integer
    Dim nm As Name
    
    'Count how many rows are needed to save all range relations
    For intI = 1 To 4
        strCtl = Choose(intI, "cbodbManySide1", _
                              "cbodbManySide2", _
                              "cbodbManySide3", _
                              "cbodbManySide4")
        
        strRange = Me(strCtl)
        If Len(strRange) Then
            Set rg = Range(Me(strCtl))
            For Each rgArea In rg.Areas
                If rgArea.Rows.Count > intMaxRows Then
                    intMaxRows = rgArea.Rows.Count
                End If
                For intJ = 1 To rgArea.Columns.Count
                    If rgArea.Cells(1, intJ).MergeCells Then
                        intJ = intJ + rgArea.Cells(1, intJ).MergeArea.Columns.Count - 1
                    End If
                    intMaxCols = intMaxCols + 1
                Next intJ
            Next
            'Add an extra row to separate each "many-side" relation
            intNumRows = intNumRows + intMaxRows + 1
            'Update columns count
            If intMaxCols > intNumCols Then
                intNumCols = intMaxCols
            End If
            intMaxRows = 0
            intMaxCols = 0
        End If
    Next intI
    Me.txtdbManySideRowsCount = intNumRows
    Me.txtdbManySideColumnsCount = intNumCols
End Function

Private Sub txtdbManySideRowsCount_Change()
    Me.txtdbManySideRowsCount1 = Me.txtdbManySideRowsCount
End Sub

Private Sub txtdbManySideColumnsCount_Change()
    Me.txtdbManySideColumnsCount1 = Me.txtdbManySideColumnsCount
End Sub

Private Sub cmdClear1_Click()
    If Len(Me.cbodbManySide2) Then
        Me.cbodbManySide1 = Me.cbodbManySide2
        Call cmdClear2_Click
    Else
        Me.cbodbManySide1 = ""
    End If
End Sub

Private Sub cmdClear2_Click()
    If Len(Me.cbodbManySide3) Then
        Me.cbodbManySide2 = Me.cbodbManySide3
        Call cmdClear3_Click
    Else
        Me.cbodbManySide2 = ""
    End If
End Sub

Private Sub cmdClear3_Click()
    If Len(Me.cbodbManySide4) Then
        Me.cbodbManySide3 = Me.cbodbManySide4
        Me.cbodbManySide4 = ""
    Else
        Me.cbodbManySide3 = ""
    End If
End Sub

Private Sub cmdClear4_Click()
    Me.cbodbManySide4 = ""
End Sub

Private Sub cmdRecordPosition_Click()
    Dim strRange As String
    Dim strMsg As String
    Dim strRecordPosition As String
        
    On Error Resume Next
    
    strRecordPosition = SetRecordPosition()
    strMsg = "Select cell to receive Record Position indicator:"
    Me.Hide
    strRange = GetRange(strMsg, "Record Position cell?", Me.txtRecordPosition)
    If Len(strRange) Then
        If Range(strRange).Column < Range(strRecordPosition).Column Then
            MsgBox "There is no room to create data navigation controls on selected cell.", vbCritical, "Invalid selection!"
        Else
            Me.txtRecordPosition = strRange
        End If
    End If
    Me.Show
End Sub

Public Function SetRecordPosition() As String
    Dim rg As Range
    Dim sngWidth As Single
    Const conNavigationButtonWidth = 17.3
    Const conRecordPosiciontCellWidth = 50.25
    
    Set rg = Range("$B$" & Me.txtdbRecordsFirstRow - 2)
    sngWidth = rg.Offset(0, -1).Width
    Do While (sngWidth < conNavigationButtonWidth * 2) And _
             (rg.Width < conRecordPosiciontCellWidth)
        sngWidth = sngWidth + rg.Width
        Set rg = rg.Offset(, 1)
    Loop
    SetRecordPosition = rg.Address(True, True)
End Function

Private Sub cmdCancel_Click()
    Dim strMsg As String
    Dim strTitle As String
    
    If Me.cmdCancel.Caption = "Remove" Then
        strMsg = "Do you really want to remove this Database structure?" & vbCrLf & vbCrLf
        strMsg = strMsg & "    Just database properties will be removed. " & vbCrLf
        strMsg = strMsg & "    Existing records will remain on the worksheet." & vbCrLf & vbCrLf
        strMsg = strMsg & "(This operation can be undone if close the workbook without saving it!)"
        strTitle = "Delete Database Properties?"
        If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbCritical, strTitle) = vbYes Then
            'Remove Database properties
            ActiveSheet.Unprotect
            Call SetDataBase(RemoveDatabase)
        End If
    End If
    Unload Me
End Sub

Private Function GetdbManySide(intRelation As Integer) As String
    Dim strMsg As String
    Dim strRange As String
        
    Me.Hide
    strMsg = "Select all column cells that belongs to the  " & intRelation & " of the 'many side' worksheet record." & vbCrLf
    strRange = GetRange(strMsg, "Many-side record cells:  " & intRelation, Me("cbodbManySide" & intRelation))
    If Len(strRange) Then
        Me("cbodbManySide" & intRelation) = strRange
    End If
    Me.Show
End Function

Private Sub cmdDefine_Click()
    Dim rg As Range
    Dim strRange As String
    Dim intI As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    ActiveSheet.Unprotect
    If Me.cmdDefine.Caption = "Define" Then
        'Define database structure
        Call SetDataBase(CreateDatabase)
    End If
    
    'Hide or show database rows
    ActiveSheet.Range(Cells(Me.txtdbRecordsFirstRow, 1), _
                      Cells(ActiveSheet.Rows.Count, 1)).EntireRow.Hidden = Me.chkHideDatabaseRows
    If Me.chkHideDatabaseRows Then
        'Unlock worksheet record cells
        Range(Me.txtdbDataValidationList).MergeArea.Locked = False
        For intI = 1 To 5
            strRange = Choose(intI, "cbodbOneSide", _
                                    "cbodbManySide1", _
                                    "cbodbManySide2", _
                                    "cbodbManySide3", _
                                    "cbodbManySide4")
            If Len(Me(strRange)) Then
                For Each rg In Range(Me(strRange)).Areas
                    For intRow = 1 To rg.Rows.Count
                        For intCol = 1 To rg.Columns.Count
                            rg.Cells(intRow, intCol).MergeArea.Locked = False
                        Next
                    Next
                Next
            End If
        Next
        'Active worksheet protection, selecting just unlocked cells
        ActiveSheet.Protect
        ActiveSheet.EnableSelection = xlUnlockedCells
    End If
    
    Unload Me
End Sub

Private Sub SetDataBase(Operation As DataBaseOperation)
    Dim nm As Name
    Dim strNameScope As String
    Dim strName As String
    Dim intRow As Integer
    Dim intI As Integer
    Const conCol = "=$B$"
    Const conColD = 4
    
    Application.ScreenUpdating = False
    intRow = Me.txtdbRecordsFirstRow
    strNameScope = "'" & ActiveSheet.Name & "'!"
    'Create database range names on columns A:B
    For intI = 0 To 14
        strName = Choose(intI + 1, "dbRecordName", _
                                   "dbDataValidationList", _
                                   "dbSavedRecords", _
                                   "dbRecordsFirstRow", _
                                   "dbOneSide", _
                                   "dbOneSideColumnsCount", _
                                   "dbManySide1", _
                                   "dbManySide2", _
                                   "dbManySide3", _
                                   "dbManySide4", _
                                   "dbManySideFirstColumn", _
                                   "dbManySideColumnsCount", _
                                   "dbManySideRowsCount", _
                                   "dbManySidePrefix", _
                                   "dbRangeOffset")
        If Operation = CreateDatabase Then
            Set nm = Names.Add(strNameScope & strName, conCol & intRow + intI, False)
            Cells(intRow + intI, 1) = strName
            Cells(intRow + intI, 2) = Me("txt" & strName)
        Else
            Set nm = Names(strNameScope & strName)
            nm.Delete
            Cells(intRow + intI, 1).ClearContents
            Cells(intRow + intI, 2).ClearContents
        End If
    Next
    
    If Operation = CreateDatabase Then
        'Define SavedRecords range name on column D
        Set nm = Names.Add(strNameScope & Me.txtdbSavedRecords, "=" & Cells(intRow, conColD).Address, False)
        'Define SavedRecords data validation list
        Range(strNameScope & Me.txtdbSavedRecords) = "New " & Me.txtdbRecordName
        Range(Me.txtdbDataValidationList).Validation.Delete
        Range(Me.txtdbDataValidationList).Validation.Add xlValidateList, , , "=" & Me.txtdbSavedRecords
        Range(Me.txtdbDataValidationList).HorizontalAlignment = xlLeft
        Range(Me.txtdbDataValidationList) = "New " & Me.txtdbRecordName
        Call CreateDatabaseButtons
    Else
        Set nm = Names(strNameScope & Me.txtdbSavedRecords)
        nm.Delete
        Range(Me.txtdbDataValidationList).Validation.Delete
        Call DeleteDatabaseButtons
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub CreateDatabaseButtons()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim rg As Range
    Dim dobjClipboard As New DataObject
    Dim strMsg As String
    Dim lngLeft As Long
    Const conColorLighBlue = 12419407
    Const conMoveButtonWidth = 17.25
    
    Set ws = Application.ActiveSheet
    
    If Me.chkCreateControlButtons Then
        'Create Database ControlButtons at right of Data Validation list
        '---------------------------------------------------------------
        Set rg = Range(Me.txtdbDataValidationList)
        If rg.MergeCells Then
            'Range has merged cells. Position on last right cell
            Set rg = Cells(rg.Row, rg.Column + rg.MergeArea.Columns.Count - 1)
        End If
        
        'Create New button
        lngLeft = rg.Left + rg.Width + 16
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, 30, rg.Height)
        shp.OnAction = "MoveNew"
        shp.OLEFormat.Object.Text = "New"
        
        'Create Save button
        lngLeft = shp.Left + shp.Width + 5
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, 30, rg.Height)
        shp.OnAction = "Save"
        shp.OLEFormat.Object.Text = "Save"
    
        'Create Delete button
        lngLeft = shp.Left + shp.Width + 5
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, 35, rg.Height)
        shp.OnAction = "Delete"
        shp.OLEFormat.Object.Text = "Delete"
    End If
    
    If Me.chkCreateNavigationButtons Then
        'Create Data Navigation buttons
        '------------------------------------------
        Set rg = Range(Me.txtRecordPosition)
        rg.Formula = "=RecordPosition()"
        rg.HorizontalAlignment = xlCenter
        rg.Font.Size = 9
        rg.Borders.LineStyle = xlContinuous
        rg.Borders.Color = conColorLighBlue
        
        'Create MoveFirst button
        lngLeft = rg.Left - 2 * conMoveButtonWidth
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, conMoveButtonWidth, rg.Height)
        shp.OnAction = "MoveFirst"
        'shp.OnAction = ws.CodeName & ".MoveFirst"
        shp.OLEFormat.Object.Text = "|<"
        
        'Create MoveFirst button
        lngLeft = rg.Left - conMoveButtonWidth
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, conMoveButtonWidth, rg.Height)
        shp.OnAction = "MovePrevious"
        shp.OLEFormat.Object.Text = "<"
        
        'Create MoveFirst button
        lngLeft = rg.Left + rg.Width
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, conMoveButtonWidth, rg.Height)
        shp.OnAction = "MoveNext"
        shp.OLEFormat.Object.Text = ">"
        
        'Create MoveFirst button
        lngLeft = rg.Left + rg.Width + conMoveButtonWidth
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, conMoveButtonWidth, rg.Height)
        shp.OnAction = "MoveLast"
        shp.OLEFormat.Object.Text = ">|"
        
        'Create MoveFirst button
        lngLeft = rg.Left + rg.Width + 2 * conMoveButtonWidth
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, lngLeft, rg.Top, conMoveButtonWidth, rg.Height)
        shp.OnAction = "MoveNew"
        shp.OLEFormat.Object.Text = "*"
    End If
    
    If Me.chkCreateControlButtons Or Me.chkCreateNavigationButtons Then
        'Copy sheet modulce code and basControlButtons code
        With dobjClipboard
            .SetText Me.txtButtonsCode.Text
            .PutInClipboard
            'Warn the user how to paste button codes on sheet module
            strMsg = "To create the database buttons code, select the worksheet code module "
            strMsg = strMsg & "place the text cursor behind the 'Option Explicit' instruction "
            strMsg = strMsg & "and press Ctrl+V to paste!"
            MsgBox strMsg, vbInformation, "WANING: How to create buttons code!"
        End With
    End If
End Sub

Public Sub DeleteDatabaseButtons()
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoFormControl Then
            If shp.FormControlType = xlButtonControl Then
                shp.Delete
            End If
        End If
    Next
End Sub


