Attribute VB_Name = "FileSaveAsDirect"
'---------------------------------------------------------------------------------------------------
' FileSaveAsDirect.bas (Robert Schwenn)
' 
' Das Makro "Optional.FileSaveAsDirect.FileSaveAsDialog" startet den klassischen Dialog "Datei Speichern als".
' Diesem Makro kann mit den anderen beiden Makros ***FileSaveAsShortcut()
' das Tastenkürzel "STRG+UMSCHALT+s" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit


' Tastenkürzel "STRG+UMSCHALT+s" wird dem Makro "FileSaveAsDialog" zugewiesen. 
Sub AssignFileSaveAsShortcut()
  Application.OnKey "+^s", "FileSaveAsDialog"
End Sub

' Tastenkürzel "STRG+UMSCHALT+s" wird auf Standard (nichts) zurückgesetzt. 
Sub ResetFileSaveAsShortcut()
  Application.OnKey "+^s"
End Sub


' Startet den klassischen Dialog "Datei Speichern als" oder einen eigenen
' Dialog "Datei Speichern als XLSM", falls die Datei noch nie gespeichert wurde
' und das Workbook-Objekt die Eigenschaft "IsWorkbookWithMacros" bietet, welche
' "True" zurückgibt.
Sub FileSaveAsDialog()
    Dim FileSaveDialog          As FileDialog
    Dim IsNewWorkbookWithMacros As Boolean
    Dim WorkbookHasVBProject    As Boolean
    Dim FilterIndex             As Long
    
    If (Not Application.ActiveWorkbook Is Nothing) Then
        
        ' Eigene Eigenschaft abfragen (möglicherweise überflüssig wegen Workbook.HasVBProject)
        IsNewWorkbookWithMacros = False
        If (Application.ActiveWorkbook.Path = "") Then
            ' New workbook (not saved yet).
            On Error Resume Next
            ' If the Workbook template provides a property "IsWorkbookWithMacros":
            IsNewWorkbookWithMacros = Application.ActiveWorkbook.IsWorkbookWithMacros
            On Error GoTo 0
        End If
        
        ' Arbeitsmappe mit Makros?
        WorkbookHasVBProject = (Application.ActiveWorkbook.HasVBProject Or IsNewWorkbookWithMacros)
        
        ' Dateifilter.
        If (WorkbookHasVBProject) Then
            FilterIndex = 2
        Else
            FilterIndex = 1
        End If
        
        ' Dateidialog.
        Set FileSaveDialog = Application.FileDialog(msoFileDialogSaveAs)
        FileSaveDialog.FilterIndex = FilterIndex
        If (FileSaveDialog.Show) Then
            ' Dialog nicht abgebrochen.
            FileSaveDialog.Execute
        End If
    End If
End Sub


' Speichert die aktive Arbeitsmappe als XLSM nach Dateidialog mit nur dieser Typwahl.
' DEPRECATED.
Private Sub SaveAsXLSM()
    
    Dim Title           As String
    Dim InitialFileName As String
    Dim FilePath        As String
    Dim FileFilter      As String
    Dim FileFilterIndex As Long
    
    InitialFileName = Application.ActiveWorkbook.Name
    Title = "Speichern unter (Neue Datei aus Vorlage mit Makros)"
    FileFilter = "Excel-Arbeitsmappe mit Makros (*.xlsm),*.xlsm"
    FileFilterIndex = 1
    FilePath = Application.GetSaveAsFilename(InitialFileName, FileFilter, FileFilterIndex, Title)
    
    If ((FilePath <> "False") And (FilePath <> "Falsch")) Then
        On Error Resume Next
        Application.ActiveWorkbook.SaveAs FileFormat:=xlOpenXMLWorkbookMacroEnabled, Filename:=FilePath, AddToMru:=True
        If (Err) Then
          ' Zielformat und Extension in jeder Hinsicht unsicher => Speichern dem Nutzer überlassen!
          MsgBox "Neu erstellte Datei konnte nicht gespeichert werden als '" & FilePath & "'." & vbNewLine & vbNewLine & " => Ist eine Datei gleichen Namens vielleicht bereits geöffnet?"
        End If
        On Error GoTo 0
    End If
End Sub

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
