Attribute VB_Name = "Globals"
'===============================================================================
'Modul Globals                                                                  
'===============================================================================

Option Explicit


Public Const AddInVersion   As String = "2.1"

' Standard-Einstellungen
Public Const EnableConditionalFormatDefault As Boolean = False
Public Const EnableFileNewDirectDefault     As Boolean = True
Public Const EnableSyncWorkDirDefault       As Boolean = True


Public oRibbon As IRibbonUI


' Region "Referenz auf das Ribbon-Objekt"
    
    ' Siehe http://social.msdn.microsoft.com/Forums/office/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
    
    ' Initialisierung der RibbonUI: Speichern einer Referenz auf das Ribbon-Objekt
    ' und als Backup eines entsprechenden Integer-Zeigers in die Add-In-inerne Tabelle.
    Public Sub OnRibbonLoad(ribbon As IRibbonUI)
        Set oRibbon = ribbon
        tabHooks.Range("A1").Value = ObjPtr(ribbon)
    End Sub
    
    ' Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    Function getRibbon() As IRibbonUI
        
        If (oRibbon Is Nothing) Then
            Dim ribbonPointer As Long
            ribbonPointer = tabHooks.Range("A1").Value
            Call CopyMemory(oRibbon, ribbonPointer, 4)
        End If
        
        Set getRibbon = oRibbon
    End Function
    
' End Region

' Region "Settings Dialog"
    
    ' Knopf "Einstellungen" gedrückt.
    Sub SettingsButtonAction(ByVal control As IRibbonControl)
        Call ShowSettingsDialog
    End Sub
    
    ' Knopf "Einstellungen" gedrückt.
    Sub ShowSettingsDialog()
        Dim Dialog As New frmSettings
        Dialog.Show
    End Sub
    
' End Region


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
