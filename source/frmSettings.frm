VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Einstellungen spezial"
   ClientHeight    =   2050
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   5050
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Modul frmSettings (Einstellungsdialog)
'===============================================================================

Option Explicit


Private Sub UserForm_Initialize()
    Me.Caption = "Einstellungen spezial (" & AddInVersion & ")"
    chkEnableConditionalFormat.Value    = ThisWorkbook.EnableConditionalFormat
    chkEnableFileNewDirect.Value        = ThisWorkbook.EnableFileNewDirect
    chkEnableSyncWorkDir.Value          = ThisWorkbook.EnableSyncWorkDir
End Sub

Private Sub btnOK_Click()
    ThisWorkbook.EnableConditionalFormat    = chkEnableConditionalFormat.Value
    ThisWorkbook.EnableFileNewDirect        = chkEnableFileNewDirect.Value
    ThisWorkbook.EnableSyncWorkDir          = chkEnableSyncWorkDir.Value
    
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
