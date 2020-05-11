Attribute VB_Name = "SaveAsPDF"
'---------------------------------------------------------------------------------------------------
' SaveAsPDF.bas (Robert Schwenn)
' 
' Das Makro "Hooks.SaveAsPDF.SaveAsPDFDialog" startet den Dialog "Als PDF ver�ffentlichen".
' Diesem Makro kann mit den anderen beiden Makros ***SaveAsPDFShortcut()
' das Tastenk�rzel "STRG+UMSCHALT+p" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit


' Startet Dialog "Als PDF ver�ffentlichen".
Sub SaveAsPDFDialog()
    On Error Resume Next
    Application.CommandBars.ExecuteMso "FileSaveAsPdfOrXps"
    On Error Goto 0
    
    ' Alternative: PDF-Export ohne jede Nachfrage (�berschreibt vorhandene PDF)
    'ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF,  _
        'Filename:="X:\Quellen\VBA\Excel\Hooks\source\Hooks_Test.pdf",  _
        'Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

' Tastenk�rzel "STRG+UMSCHALT+p" wird dem Makro "SaveAsPDFDialog" zugewiesen. 
Sub AssignSaveAsPDFShortcut()
  Application.OnKey "+^p", "SaveAsPDFDialog"
End Sub

' Tastenk�rzel "STRG+UMSCHALT+p" wird auf Standard zur�ckgesetzt. 
Sub ResetSaveAsPDFShortcut()
  Application.OnKey "+^p"
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
