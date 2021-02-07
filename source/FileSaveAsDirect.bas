Attribute VB_Name = "FileSaveAsDirect"
'---------------------------------------------------------------------------------------------------
' FileSaveAsDirect.bas (Robert Schwenn)
' 
' Das Makro "Optional.FileSaveAsDirect.FileSaveAsDialog" startet den klassischen Dialog "Datei Speichern als".
' Diesem Makro kann mit den anderen beiden Makros ***FileSaveAsShortcut()
' das Tastenkürzel "STRG+UMSCHALT+s" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit

'Private Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long


' Tastenkürzel "STRG+UMSCHALT+s" wird dem Makro "FileSaveAsDialog" zugewiesen. 
Sub AssignFileSaveAsShortcut()
  Application.OnKey "+^s", "FileSaveAsDialog"
End Sub

' Tastenkürzel "STRG+UMSCHALT+s" wird auf Standard (nichts) zurückgesetzt. 
Sub ResetFileSaveAsShortcut()
  Application.OnKey "+^s"
End Sub


' Startet den klassischen Dialog "Datei Speichern als".
Sub FileSaveAsDialog()
    'On Error Resume Next
    'Dim WorkbookDir As String
    
    ' Arbeitsverzeichnis setzen auf das der aktiven Arbeitsmappe.
    'If (Not ActiveWorkbook Is Nothing) Then
     '   WorkbookDir = ActiveWorkbook.Path
     '   If (Not (WorkbookDir = "")) Then
     '       ' Zunächst Sonderfall abfangen (Wurzelverzeichnis).
     '       If (Right(WorkbookDir, 1) = ":") Then
     '           WorkbookDir = WorkbookDir & Application.PathSeparator
     '       End If
     '       ' Arbeitsverzeichnis ändern.
     '       SetCurrentDirectory WorkbookDir
     '   End If
    'End If
    
    ' Dateidialog.
    Application.CommandBars.ExecuteMso "FileSaveAs"
End Sub


'' Setzt das Arbeitsverzeichnis (auch für UNC-Pfade).
'Private Sub SetCurrentDirectory(Path As String)
'    Dim lReturn As Long
'    lReturn = SetCurrentDirectoryA(Path)
'    'If lReturn = 0 Then 
'    '    MsgBox "Error setting path"
'    'End If
'End Sub

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
