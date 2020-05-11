Attribute VB_Name = "FileOpenDirect"
'---------------------------------------------------------------------------------------------------
' FileOpenDirect.bas (Robert Schwenn)
' 
' Das Makro "Hooks.FileOpenDirect.FileOpenDialog" startet den klassischen Dialog "Datei Öffnen".
' Diesem Makro kann mit den anderen beiden Makros ***FileOpenShortcut()
' das Tastenkürzel "STRG+UMSCHALT+o" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit


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
    Application.CommandBars.ExecuteMso "FileOpen"
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
