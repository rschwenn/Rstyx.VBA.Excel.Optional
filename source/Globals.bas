Attribute VB_Name = "Globals"
'===============================================================================
' Modul Globals                                                                  
'===============================================================================

Option Explicit


Public Const AddInVersion   As String = "2.4"

' Standard-Einstellungen
Public Const EnableConditionalFormatDefault As Boolean = False
Public Const EnableFileNewDirectDefault     As Boolean = True
Public Const EnableSyncWorkDirDefault       As Boolean = True
Public Const EnableSaveAsPDFDefault         As Boolean = True


Private oRibbon As IRibbonUI


' Region "Referenz auf das Ribbon-Objekt"
    
    ' Siehe http://social.msdn.microsoft.com/Forums/office/en-US/99a3f3af-678f-4338-b5a1-b79d3463fb0b/how-to-get-the-reference-to-the-iribbonui-in-vba
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
    
    ' Initialisierung der RibbonUI: Speichern einer Referenz auf das Ribbon-Objekt
    ' und als Backup eines entsprechenden Integer-Zeigers in die Add-In-interne Tabelle.
    Public Sub OnHooksRibbonLoad(ribbon As IRibbonUI)
        Set oRibbon = ribbon
        tabHooks.Range("A1").Value = ObjPtr(ribbon)
    End Sub
    
    ' Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    Function getHooksRibbon() As IRibbonUI
        ' "oRibbon" ist normalerweise nur dann "Nothing", wenn das AddIn wegen eines Fehlers gestoppt wurde.
        ' Dann kann der vorher gespeicherte Zeiger verwendet werden.
        ' ABER: Wenn das AddIn nicht schreibgeschützt ist, kann der Zeiger auch veraltet sein.
        '       => Dann stürzt Excel ab und nichts geht mehr.
        If (oRibbon Is Nothing) Then
            If (ThisWorkbook.ReadOnly) Then
                Dim ribbonPointer As Long
                ribbonPointer = tabHooks.Range("A1").value
                If (ribbonPointer > 0) Then
                    On Error Resume Next  ' Nützt nix!
                    Call CopyMemory(oRibbon, ribbonPointer, 4)
                    On Error GoTo 0
                End If
            End If
        End If
        
        Set getHooksRibbon = oRibbon
    End Function
    
    ' Status-Aktualisierung aller Ribbon-Steuerelemente erzwingen.
    ' Falls das AddIn gestoopt wurde, impliziert diese Routine auch dessen Neustart
    ' durch die Verwendung der Eigenschaft ThisWorkbook.AktiveTabelle...
    Public Sub UpdateHooksRibbon()
        On Error Resume Next
        getHooksRibbon().Invalidate
        On Error Goto 0
    End Sub
    
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
