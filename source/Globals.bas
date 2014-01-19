Attribute VB_Name = "Globals"
'===============================================================================
'Modul Globals                                                                  
'===============================================================================


Option Explicit


' Standard-Einstellungen
Public Const EnableFileNewDirectDefault As Boolean = True
Public Const EnableSyncWorkDirDefault   As Boolean = True



' Region "Settings Dialog"
    
    ' Knopf "Einstellungen" gedrückt.
    Sub SettingsButtonAction(ByVal control As IRibbonControl)
        Call ShowSettingsDialog
    End Sub
    
    ' Knopf "Einstellungen" gedrückt.
    Sub ShowSettingsDialog()
        Dim Dialog As New frmSettings
        Dialog.Show()
    End Sub
    
' End Region


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
