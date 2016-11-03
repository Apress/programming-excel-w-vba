VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApplicationEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set app = Application
End Sub

Private Sub Class_Terminate()
    Set app = Nothing
End Sub

Private Sub app_AfterCalculate()
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application AfterCalculate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_AfterCalculate event fired" & vbCrLf
    strMsg = strMsg & "One or more workbooks was calculated!"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_NewWorkbook(ByVal Wb As Workbook)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application NewWorkbook event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_NewWorkbook event fired." & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "(Excel has opened another workbook!)"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_SheetActivate(ByVal Sh As Object)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetActivate event fired"
    strMsg = strMsg & " Application_SheetActivate event fired" & vbCrLf
    strMsg = "CApplicationEvents object" & vbCrLf
    strMsg = strMsg & "Applicatrion_SheetActivate event fired"
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & "Sheet activated is " & Sh.Name
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetBeforeDoubleClick event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & " Application_SheetBeforeDoubleClick event fired." & vbCrLf
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & "Sheet double clicked is: " & Sh.Name & vbCrLf
    strMsg = strMsg & "Do you want to double click cell " _
             & Target.Address & " from " & Sh.Name & " and put it on Edit mode?"
    If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub app_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetBeforeRightClick event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_SheetBeforeRightClick event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & " Sheet Right Clicked is " & Sh.Name & vbCrLf
    strMsg = strMsg & "Do you want to right click cell " _
             & Target.Address & " from " & Sh.Name & " and show the Context Menu?"
    If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub app_SheetCalculate(ByVal Sh As Object)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetCalculate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_SheetCalculate event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & "Sheet " & Sh.Name & " was calculated!"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetChange event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_SheetChange event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & "Sheet changed is " & Sh.Name & vbCrLf
    strMsg = strMsg & "Cell(s) changed is(are) " & Target.Address
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_SheetDeactivate(ByVal Sh As Object)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetDeactivate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_SheetDeactivate event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & "Sheet Deactivated is " & Sh.Name
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application SheetSelectionChange( event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_SheetSelectionChange event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Sh.Parent.Name & vbCrLf
    strMsg = strMsg & "Sheet where selection changed is " & Sh.Name & vbCrLf
    strMsg = strMsg & "Cell(s) selected is(are) " & Target.Address
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WindowActivate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WindowActivate event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "(Excel has more than one workbook opened or the workbook is opening)"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WindowDeactivate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WindowDeactivate event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "(Excel has more than one workbook opened or the workbook is closing)"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WindowResize event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WindowResize event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "(The worbook window had been restored, maximized or resized inside Excel)"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WorkbookActivate(ByVal Wb As Workbook)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookActivate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookActivate event fired." & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "(Excel has more than one workbook opened or the workbook is opening!)"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookAfterSave event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookAfterSave event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "The saving process was " & IIf(Success, "successfull", "failed")
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookBeforeClose event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookBeforeClose event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "Do you want to close the workbook?"
    If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub app_WorkbookBeforePrint(ByVal Wb As Workbook, Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookBeforePrint event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookBeforePrint event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "Do you want to print the workbook?"
    If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub app_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookBeforeSave event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookBeforeSave event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "Do you want to save the workbook?"
    If MsgBox(strMsg, vbYesNo + vbQuestion, strTitle) = vbNo Then
        Cancel = True
    Else
        If SaveAsUI Then
            MsgBox "Save As dialog box will now be opened: Folder or Workbook name can be changed!"
        End If
    End If
End Sub

Private Sub app_WorkbookDeactivate(ByVal Wb As Workbook)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookDeactivate event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookDeactivate event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "(Excel has more than one workbook opened or the workbook is closing!)"
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WorkbookNewChart(ByVal Wb As Workbook, ByVal Ch As Chart)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookNewChart event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookNewChart event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "Chart inserted is" & Ch.Name & vbCrLf
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookNewSheet event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookNewSheet event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    strMsg = strMsg & "Sheet inserted is" & Sh.Name & vbCrLf
    MsgBox strMsg, vbInformation, strTitle
End Sub

Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
    Dim strMsg As String
    Dim strTitle As String
    
    strTitle = "Application WorkbookOpen event fired"
    strMsg = "CApplicationEvents object" & vbCrLf & vbCrLf
    strMsg = strMsg & "Application_WorkbookOpen event fired" & vbCrLf
    strMsg = strMsg & "Workbook: " & Wb.Name & vbCrLf
    MsgBox strMsg, vbInformation, strTitle
End Sub

