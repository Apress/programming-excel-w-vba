VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetnameChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event SheetNameChange(ByVal Sh As Object, ByRef Cancel As Boolean)

Private WithEvents app As Excel.Application
Attribute app.VB_VarHelpID = -1
Private mWks As Worksheet
Private mstrWksName As String

Private Sub Class_Initialize()
    Set app = Application
    Set mWks = ActiveSheet
    mstrWksName = mWks.Name
End Sub

Private Sub Class_Terminate()
    Set mWks = Nothing
    Set app = Nothing
End Sub

Public Sub NameChange(Optional wks As Worksheet)
    Dim bolCancel As Boolean
    
    If wks Is Nothing Then
        Set wks = app.ActiveSheet
    End If
    
    If mstrWksName <> mWks.Name Then
        RaiseEvent SheetNameChange(mWks, bolCancel)
        If bolCancel Then
            mWks.Name = mstrWksName
        End If
    End If

    Set mWks = wks
    mstrWksName = mWks.Name
End Sub

Private Sub app_SheetActivate(ByVal Sh As Object)
    Call NameChange(Sh)
End Sub

Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Call NameChange(Sh)
End Sub

Private Sub app_SheetDeactivate(ByVal Sh As Object)
    Call NameChange(Sh)
End Sub

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Call NameChange(Sh)
End Sub

Private Sub app_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call NameChange
End Sub
