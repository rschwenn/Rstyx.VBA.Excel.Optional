Attribute VB_Name = "FileOpenDirect"
'---------------------------------------------------------------------------------------------------
' FileOpenDirect.bas (Robert Schwenn)
' 
' Das Makro "Optional.FileOpenDirect.FileOpenDialog" startet den klassischen Dialog "Datei Öffnen".
' Diesem Makro kann mit den anderen beiden Makros ***FileOpenShortcut()
' das Tastenkürzel "STRG+UMSCHALT+o" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit

Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long


' Tastenkürzel "STRG+UMSCHALT+o" wird dem Makro "FileOpenDialog" zugewiesen. 
Sub AssignFileOpenShortcut()
  Application.OnKey "+^o", "FileOpenDialog"
End Sub

' Tastenkürzel "STRG+UMSCHALT+o" wird auf Standard (nichts) zurückgesetzt. 
Sub ResetFileOpenShortcut()
  Application.OnKey "+^o"
End Sub


' Startet den klassischen Dialog "Datei Neu".
Sub FileOpenDialog()
    'On Error Resume Next
    Dim WorkbookDir As String
    
    ' Arbeitsverzeichnis setzen auf das der aktiven Arbeitsmappe.
    If (Not ActiveWorkbook Is Nothing) Then
        WorkbookDir = ActiveWorkbook.Path
        If (Not (WorkbookDir = "")) Then
            ' Zunächst Sonderfall abfangen (Wurzelverzeichnis).
            If (Right(WorkbookDir, 1) = ":") Then
                WorkbookDir = WorkbookDir & Application.PathSeparator
            End If
            ' Arbeitsverzeichnis ändern.
            SetCurrentDirectory WorkbookDir
        End If
    End If
    
    ' Dateidialog.
    Application.CommandBars.ExecuteMso "FileOpen"
End Sub


' Setzt das Arbeitsverzeichnis (auch für UNC-Pfade).
Private Sub SetCurrentDirectory(Path As String)
    Dim lReturn As Long
    lReturn = SetCurrentDirectoryA(Path)
    'If lReturn = 0 Then 
    '    MsgBox "Error setting path"
    'End If
End Sub

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
