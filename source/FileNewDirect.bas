Attribute VB_Name = "FileNewDirect"
'---------------------------------------------------------------------------------------------------
' FileNewDirect.bas (Robert Schwenn)
' 
' Das Makro "Hooks.FileNewDirect.FileNewDialog" startet den klassischen Dialog "Datei Neu".
' Diesem Makro kann mit den anderen beiden Makros ***FileNewShortcut()
' das Tastenkürzel "STRG+UMSCHALT+p" zugewiesen bzw. entzogen werden.
' 
' Der Back-Office-Knopf "Neu (Dialog)" wird via XML angelegt.
' Dessen Sichtbarkeit wird gesteuert via Callback "getVisibleFileNewButton()"
' das die Eigenschaft "EnableFileNewButton" zurückgibt.
'---------------------------------------------------------------------------------------------------

Option Explicit


' Ribbon-Callback
Public Sub getVisibleFileNewButton(control As IRibbonControl, ByRef visible)
    
    visible = ThisWorkbook.EnableFileNewButton
End Sub
    
' Back-Office-Knopf "Neu (Dialog)" gedrückt.
Sub FileNewButtonAction(ByVal control As IRibbonControl)
    Call FileNewDialog
End Sub


' Tastenkürzel "STRG+UMSCHALT+n" wird dem Makro "FileNewDialog" zugewiesen. 
Sub AssignFileNewShortcut()
  Application.OnKey "+^n", "FileNewDialog"
End Sub

' Tastenkürzel "STRG+UMSCHALT+n" wird auf Standard (nichts) zurückgesetzt. 
Sub ResetFileNewShortcut()
  Application.OnKey "+^n"
End Sub


' Startet den klassischen Dialog "Datei Neu".
Sub FileNewDialog()
    'On Error Resume Next
    
    ' Listenansicht aktivieren
    SendKeys "%2"
    
    ' Dialog starten.
    Application.CommandBars.ExecuteMso "FileNew"
    'Application.Dialogs(xlDialogNew).Show
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
