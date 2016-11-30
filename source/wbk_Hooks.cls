VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbk_Hooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute cbbFileNew.VB_VarHelpID = -1
'---------------------------------------------------------------------------------------------------
' wbk_Hooks.cls (Robert Schwenn)
' 
' 24.11.2008 v1.0  Der Men�punkt "Datei -> Neu..." wird hier verarbeitet:
'                    Statt den Aufgabenbereich anzuzeigen, startet direkt der
'                    Dialog "Vorlage w�hlen" (wie Excel 97) im Listenmodus
' 04.01.2014 v2.0  - Umstrukturiert: 
'                    - Je Hook ein eigenes Modul
'                    - In wbk_Hooks.cls: je Hook eine Aktivierungs- und Deaktivierungsroutine
'                      sowie eine Eigenschaft (Hauptschalter)
'                  - Anpassung f�r Excel 2010:
'                    - Hook f�r "Datei -> Neu..." umgestrickt.
'                    - Synchronisierung des Excel-Arbeitsverzeichnisses mit dem Speicherort
'                      der aktiven Arbeitsmappe, damit der �ffnen-Dialog dort startet.
'                      (Der Speichern-unter-Dialog startet ohnehin dort).
'                    - Einstellungsdialog in Register-Tab "Spezial"
' 24.01.2014 v2.1  - Neu: Kontextmen�eintrag "Bedingte Formatierung"
' 16.02.2014 v2.2  - Ribbon-Kode: Excel-weit eindeutige Routinenamen (Problemvermeidung)
' 20.11.2016 v2.3  - Bugfix: Synchronisierung des Excel-Arbeitsverzeichnisses arbeitet jetzt
'                    auch mit UNC-Pfaden (bisher trat ein Fehler auf)
' 30.11.2016 v2.4  - Neu: PDF-Export per Tastenk�rzel "STRG+UMSCHALT+p".
'---------------------------------------------------------------------------------------------------
 
Option Explicit

Private SyncWorkDir                             As SyncWorkDir
Private blnEnableSyncWorkDir                    As Boolean
Private blnEnableSyncWorkDirInitialized         As Boolean

Private blnEnableFileNewDirect                  As Boolean
Private blnEnableFileNewDirectInitialized       As Boolean

Private blnEnableConditionalFormat              As Boolean
Private blnEnableConditionalFormatInitialized   As Boolean

Private blnEnableSaveAsPDF                      As Boolean
Private blnEnableSaveAsPDFInitialized           As Boolean


Private strVBAHostNameVersion   As String
Private strAddInName            As String


' Initialisierung
Private Sub Workbook_Open()
    ' Initialisierung der Klassen-Hooks.
    'If (Me.EnableConditionalFormat) Then Call ActivateConditionalFormat
    'If (Me.EnableFileNewDirect) Then Call ActivateFileNewDirect
    If (Me.EnableSyncWorkDir) Then Call ActivateSyncWorkDir
    If (Me.EnableSaveAsPDF)   Then Call ActivateSaveAsPDF
End Sub


' Region "Add-In Umgebung"
    
    '(Toleranz gegen�ber zwischenzeitlichem Abbruch des Add-Ins)
    
    Public Property Get VBAHostNameVersion() As String
        If (strVBAHostNameVersion = "") Then
            strVBAHostNameVersion = Application.Name & " " & Application.Version
        End If
        VBAHostNameVersion = strVBAHostNameVersion
    End Property
    
    Public Property Get AddInName() As String
        If (strAddInName = "") Then
            strAddInName = ThisWorkbook.Name
        End If
        AddInName = strAddInName
    End Property
'


' Region "ConditionalFormat" (Kontextmen�eintrag "Bedingte Formatierung")
    
    Public Property Get EnableConditionalFormat() As Boolean
        If (Not blnEnableConditionalFormatInitialized) Then
            On Error Resume Next
            blnEnableConditionalFormat = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableConditionalFormat", EnableConditionalFormatDefault))
            On Error Goto 0
            blnEnableConditionalFormatInitialized = True
        End If
        EnableConditionalFormat = blnEnableConditionalFormat
    End Property
    
    Public Property Let EnableConditionalFormat(NewValue As Boolean)
        If (blnEnableConditionalFormat Xor NewValue) Then
            blnEnableConditionalFormat = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableConditionalFormat", CStr(blnEnableConditionalFormat)
            On Error Goto 0
            If (blnEnableConditionalFormat) Then
                Call ActivateConditionalFormat()
            Else
                Call DeactivateConditionalFormat()
            End If
        End If
    End Property
    
    Private Sub ActivateConditionalFormat()
        Call UpdateHooksRibbon
    End Sub
    
    Private Sub DeactivateConditionalFormat()
        Call UpdateHooksRibbon
    End Sub
'


' Region "FileNewDirect" ("Datei -> Neu..." f�hrt direkt zum Vorlagendialog)
    
    Public Property Get EnableFileNewDirect() As Boolean
        If (Not blnEnableFileNewDirectInitialized) Then
            On Error Resume Next
            blnEnableFileNewDirect = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableFileNewDirect", EnableFileNewDirectDefault))
            On Error Goto 0
            blnEnableFileNewDirectInitialized = True
        End If
        EnableFileNewDirect = blnEnableFileNewDirect
    End Property
    
    Public Property Let EnableFileNewDirect(NewValue As Boolean)
        If (blnEnableFileNewDirect Xor NewValue) Then
            blnEnableFileNewDirect = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableFileNewDirect", CStr(blnEnableFileNewDirect)
            On Error Goto 0
            If (blnEnableFileNewDirect) Then
                Call ActivateFileNewDirect()
            Else
                Call DeactivateFileNewDirect()
            End If
        End If
    End Property
    
    Private Sub ActivateFileNewDirect()
        'Set FileNewDirect = New FileNewDirect
    End Sub
    
    Private Sub DeactivateFileNewDirect()
        'Set FileNewDirect = Nothing
    End Sub
'


' Region "SyncWorkDir" (Synchronisierung Arbeitsverz. mit dem Speicherort der aktiven Arbeitsmappe)
    
    Public Property Get EnableSyncWorkDir() As Boolean
        If (Not blnEnableSyncWorkDirInitialized) Then
            On Error Resume Next
            blnEnableSyncWorkDir = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableSyncWorkDir", EnableSyncWorkDirDefault))
            On Error Goto 0
            blnEnableSyncWorkDirInitialized = True
        End If
        EnableSyncWorkDir = blnEnableSyncWorkDir
    End Property
    
    Public Property Let EnableSyncWorkDir(NewValue As Boolean)
        If (blnEnableSyncWorkDir Xor NewValue) Then
            blnEnableSyncWorkDir = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableSyncWorkDir", CStr(blnEnableSyncWorkDir)
            On Error Goto 0
            If (blnEnableSyncWorkDir) Then
                Call ActivateSyncWorkDir()
            Else
                Call DeactivateSyncWorkDir()
            End If
        End If
    End Property
    
    Private Sub ActivateSyncWorkDir()
        Set SyncWorkDir = New SyncWorkDir
    End Sub
    
    Private Sub DeactivateSyncWorkDir()
        Set SyncWorkDir = Nothing
    End Sub
'


' Region "SaveAsPDF" (Tastenk�rzel "STRG+UMSCHALT+p" startet den Dialog "Als PDF ver�ffentlichen")
    
    Public Property Get EnableSaveAsPDF() As Boolean
        If (Not blnEnableSaveAsPDFInitialized) Then
            On Error Resume Next
            blnEnableSaveAsPDF = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableSaveAsPDF", EnableSaveAsPDFDefault))
            On Error Goto 0
            blnEnableSaveAsPDFInitialized = True
        End If
        EnableSaveAsPDF = blnEnableSaveAsPDF
    End Property
    
    Public Property Let EnableSaveAsPDF(NewValue As Boolean)
        If (blnEnableSaveAsPDF Xor NewValue) Then
            blnEnableSaveAsPDF = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableSaveAsPDF", CStr(blnEnableSaveAsPDF)
            On Error Goto 0
            If (blnEnableSaveAsPDF) Then
                Call ActivateSaveAsPDF()
            Else
                Call DeactivateSaveAsPDF()
            End If
        End If
    End Property
    
    Private Sub ActivateSaveAsPDF()
        Call AssignSaveAsPDFShortcut
    End Sub
    
    Private Sub DeactivateSaveAsPDF()
        Call ResetSaveAsPDFShortcut
    End Sub
'

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4: