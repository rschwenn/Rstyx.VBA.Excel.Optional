Attribute VB_Name = "FileOpenDirect"
'---------------------------------------------------------------------------------------------------
' FileOpenDirect.bas (Robert Schwenn)
' 
' Das Makro "Optional.FileOpenDirect.FileOpenDialog" startet den klassischen Dialog "Datei Öffnen".
' Diesem Makro kann mit den anderen beiden Makros ***FileOpenShortcut()
' das Tastenkürzel "STRG+UMSCHALT+o" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit

' Tastenkürzel "STRG+UMSCHALT+o" wird dem Makro "FileOpenDialog" zugewiesen. 
Sub AssignFileOpenShortcut()
  Application.OnKey "+^o", "FileOpenDialog"
End Sub

' Tastenkürzel "STRG+UMSCHALT+o" wird auf Standard zurückgesetzt. 
Sub ResetFileOpenShortcut()
  Application.OnKey "+^o"
End Sub


' Startet den klassischen Dialog "Datei Öffnen".
Sub FileOpenDialog()
    'On Error Resume Next
    Dim FileOpenDialog As FileDialog
    
    ' Arbeitsverzeichnis setzen auf das der aktiven Arbeitsmappe.
    If (Not ActiveWorkbook Is Nothing) Then
        SetCurrentDirectory ActiveWorkbook.Path
    End If
    
    Set FileOpenDialog = Application.FileDialog(msoFileDialogOpen)
    FileOpenDialog.FilterIndex = 2
    If (FileOpenDialog.Show) Then
        ' Dialog nicht abgebrochen.
        FileOpenDialog.Execute
        
        ' Ereignisse werden hier unterdrückt (durch FileOpenDialog.Execute ?)
        ' => Für GeoTools nachhelfen durch Aufruf der dafür vorgesehenen Routine:
        On Error Resume Next
        Call Application.Run("WorkbookOpened")
        On Error Goto 0
    End If
End Sub


' Setzt das Arbeitsverzeichnis (auch für UNC-Pfade).
Private Sub SetCurrentDirectory(Path As String)

    Dim WshShell   As IWshRuntimeLibrary.WshShell
    Dim oFs        As Scripting.FileSystemObject
    Set WshShell = New IWshRuntimeLibrary.WshShell
    Set oFs      = New Scripting.FileSystemObject
    
    ' Zunächst Sonderfall abfangen (Wurzelverzeichnis).
    If (Not (Path = "")) Then
        If (Right(Path, 1) = ":") Then
            Path = Path & Application.PathSeparator
        End If
    End If
    
    ' Arbeitsverzeichnis setzen.
    If (oFs.FolderExists(Path)) Then
        WshShell.CurrentDirectory = Path
    End If
    
    Set WshShell = Nothing
    Set oFs      = Nothing
End Sub

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
