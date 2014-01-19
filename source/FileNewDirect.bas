Attribute VB_Name = "FileNewDirect"
'---------------------------------------------------------------------------------------------------
' FileNewDirect.bas (Robert Schwenn)
' 
' Das Makro "Hooks.FileNewDirect.FileNewSmart" überschreibt die Tastenkombination
' STRG+N und die Standard-Aktion "Neue, leere Datei anlegen".
' - Optional führt es genau diese Standard-Aktion aus.
' - Anderenfalls startet direkt der Dialog "Vorlage wählen" im Listenmodus.
'---------------------------------------------------------------------------------------------------

Option Explicit

' Umgeleiteter Standard-Befehl
Sub FromFileNewDefault(ByVal control As IRibbonControl, ByRef cancelDefault)
    cancelDefault = True
    Call FileNewSmart
End Sub

' Neue Datei wird erstellt.
Sub FileNewSmart()
    If (ThisWorkbook.EnableFileNewDirect) Then
        'On Error Resume Next
        SendKeys "%2"
        Application.Dialogs(xlDialogNew).Show
    Else
        Application.Workbooks.Add
    End If
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
