VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyncWorkDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------------------
' SyncWorkDir.cls (Robert Schwenn)
' 
' Eine aktive Instanz übernimmt automatisch die Synchronisierung des 
' Excel-Arbeitsverzeichnisses mit dem Speicherort der aktiven Arbeitsmappe,
' damit der Öffnen-Dialog dort startet.
' (Der Speichern-unter-Dialog startet ohnehin dort).
'---------------------------------------------------------------------------------------------------

Option Explicit

Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

Dim WithEvents App    As Application

Private Sub Class_Initialize()
    Set App = Application
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub

' Setzt das Arbeitsverzeichnis (auch für UNC-Pfade).
Private Sub SetCurrentDirectory(Path As String)
    Dim lReturn As Long
    lReturn = SetCurrentDirectoryA(Path)
    'If lReturn = 0 Then 
    '    MsgBox "Error setting path"
    'End If
End Sub

' Reaktion auf das Aktivieren einer Arbeitsmappe.
Private Sub App_WorkbookActivate(ByVal Wb As Excel.Workbook)
    'On Error Resume Next
    Dim WorkbookDir As String
    WorkbookDir = Wb.Path
    
    If (Not (WorkbookDir = "")) Then
        ' Zunächst Sonderfall abfangen (Wurzelverzeichnis).
        If (Right(WorkbookDir, 1) = ":") Then
            WorkbookDir = WorkbookDir & Application.PathSeparator
        End If
        ' Arbeitsverzeichnis ändern.
        'ChDrive WorkbookDir
        'ChDir WorkbookDir
        SetCurrentDirectory WorkbookDir
    End If
    On Error GoTo 0
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
