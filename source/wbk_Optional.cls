VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbk_Optional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute cbbFileNew.VB_VarHelpID = -1
'---------------------------------------------------------------------------------------------------
' wbk_Optional.cls (Robert Schwenn)
' 
' 24.11.2008 v1.0  Der Menüpunkt "Datei -> Neu..." wird hier verarbeitet:
'                    Statt den Aufgabenbereich anzuzeigen, startet direkt der
'                    Dialog "Vorlage wählen" (wie Excel 97) im Listenmodus
' 04.01.2014 v2.0  - Umstrukturiert: 
'                    - Je Hook ein eigenes Modul
'                    - In wbk_Hooks.cls: je Hook eine Aktivierungs- und Deaktivierungsroutine
'                      sowie eine Eigenschaft (Hauptschalter)
'                  - Anpassung für Excel 2010:
'                    - Hook für "Datei -> Neu..." umgestrickt.
'                    - Synchronisierung des Excel-Arbeitsverzeichnisses mit dem Speicherort
'                      der aktiven Arbeitsmappe, damit der Öffnen-Dialog dort startet.
'                      (Der Speichern-unter-Dialog startet ohnehin dort).
'                    - Einstellungsdialog in Register-Tab "Spezial"
' 24.01.2014 v2.1  - Neu: Kontextmenüeintrag "Bedingte Formatierung"
' 16.02.2014 v2.2  - Ribbon-Kode: Excel-weit eindeutige Routinenamen (Problemvermeidung)
' 20.11.2016 v2.3  - Bugfix: Synchronisierung des Excel-Arbeitsverzeichnisses arbeitet jetzt
'                    auch mit UNC-Pfaden (bisher trat ein Fehler auf)
' 30.11.2016 v2.4  - Neu: PDF-Export per Tastenkürzel "STRG+UMSCHALT+p".
' 11.05.2020 v3.0  - Umstellung auf 64 bit und Office 2016.
'                  - Synchronisierung des Excel-Arbeitsverzeichnisses aufgegeben.
'                    (scheint nun Excel-Standard zu sein wegen SDI)
'                  - FileNewDirect ersetzt durch FileNewShortcut und FileNewButton.
'                    (Die Standardaktion "Datei -> Neu..." wird nicht mehr abgefangen)
'                  - Neu: FileNewOpenShortcut
' 21.05.2020 v3.1  - Add-In umbenannt von "Hooks" nach "Optional"
'---------------------------------------------------------------------------------------------------
 
Option Explicit

Private SyncWorkDir                             As SyncWorkDir
Private blnEnableSyncWorkDir                    As Boolean
Private blnEnableSyncWorkDirInitialized         As Boolean

Private blnEnableFileNewShortcut                As Boolean
Private blnEnableFileNewShortcutInitialized     As Boolean

Private blnEnableFileNewButton                  As Boolean
Private blnEnableFileNewButtonInitialized       As Boolean

Private blnEnableFileOpenShortcut               As Boolean
Private blnEnableFileOpenShortcutInitialized    As Boolean

Private blnEnableConditionalFormat              As Boolean
Private blnEnableConditionalFormatInitialized   As Boolean

Private blnEnableSaveAsPDF                      As Boolean
Private blnEnableSaveAsPDFInitialized           As Boolean


Private strVBAHostNameVersion   As String
Private strAddInName            As String


' Initialisierung
Private Sub Workbook_Open()
    ' Initialisierung der Klassen.
    If (Me.EnableConditionalFormat) Then Call ActivateConditionalFormat
    If (Me.EnableFileNewButton)     Then Call ActivateFileNewButton
    If (Me.EnableFileNewShortcut)   Then Call ActivateFileNewShortcut
    If (Me.EnableFileOpenShortcut)  Then Call ActivateFileOpenShortcut
    If (Me.EnableSaveAsPDF)         Then Call ActivateSaveAsPDF
    'If (Me.EnableSyncWorkDir)       Then Call ActivateSyncWorkDir
End Sub


' Region "Add-In Umgebung"
    
    '(Toleranz gegenüber zwischenzeitlichem Abbruch des Add-Ins)
    
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


' Region "ConditionalFormat" (Kontextmenüeintrag "Bedingte Formatierung")
    
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
        Call UpdateOptionalRibbon
    End Sub
    
    Private Sub DeactivateConditionalFormat()
        Call UpdateOptionalRibbon
    End Sub
'


' Region "FileNewDialog" (klassischer Vorlagendialog)
    
    ' Shortcut
    
    Public Property Get EnableFileNewShortcut() As Boolean
        If (Not blnEnableFileNewShortcutInitialized) Then
            On Error Resume Next
            blnEnableFileNewShortcut = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableFileNewShortcut", EnableFileNewShortcutDefault))
            On Error Goto 0
            blnEnableFileNewShortcutInitialized = True
        End If
        EnableFileNewShortcut = blnEnableFileNewShortcut
    End Property
    
    Public Property Let EnableFileNewShortcut(NewValue As Boolean)
        If (blnEnableFileNewShortcut Xor NewValue) Then
            blnEnableFileNewShortcut = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableFileNewShortcut", CStr(blnEnableFileNewShortcut)
            On Error Goto 0
            If (blnEnableFileNewShortcut) Then
                Call ActivateFileNewShortcut()
            Else
                Call DeactivateFileNewShortcut()
            End If
        End If
    End Property
    
    Private Sub ActivateFileNewShortcut()
        Call AssignFileNewShortcut()
    End Sub
    
    Private Sub DeactivateFileNewShortcut()
        Call ResetFileNewShortcut()
    End Sub
    
    
    ' Back Office Button.
    
    Public Property Get EnableFileNewButton() As Boolean
        If (Not blnEnableFileNewButtonInitialized) Then
            On Error Resume Next
            blnEnableFileNewButton = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableFileNewButton", EnableFileNewButtonDefault))
            On Error Goto 0
            blnEnableFileNewButtonInitialized = True
        End If
        EnableFileNewButton = blnEnableFileNewButton
    End Property
    
    Public Property Let EnableFileNewButton(NewValue As Boolean)
        If (blnEnableFileNewButton Xor NewValue) Then
            blnEnableFileNewButton = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableFileNewButton", CStr(blnEnableFileNewButton)
            On Error Goto 0
            If (blnEnableFileNewButton) Then
                Call ActivateFileNewButton()
            Else
                Call DeactivateFileNewButton()
            End If
        End If
    End Property
    
    Private Sub ActivateFileNewButton()
        Call UpdateOptionalRibbon
    End Sub
    
    Private Sub DeactivateFileNewButton()
        Call UpdateOptionalRibbon
    End Sub
'


' Region "FileOpenDialog" (klassischer Öffnen-Dialog)
    
    Public Property Get EnableFileOpenShortcut() As Boolean
        If (Not blnEnableFileOpenShortcutInitialized) Then
            On Error Resume Next
            blnEnableFileOpenShortcut = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableFileOpenShortcut", EnableFileOpenShortcutDefault))
            On Error Goto 0
            blnEnableFileOpenShortcutInitialized = True
        End If
        EnableFileOpenShortcut = blnEnableFileOpenShortcut
    End Property
    
    Public Property Let EnableFileOpenShortcut(NewValue As Boolean)
        If (blnEnableFileOpenShortcut Xor NewValue) Then
            blnEnableFileOpenShortcut = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableFileOpenShortcut", CStr(blnEnableFileOpenShortcut)
            On Error Goto 0
            If (blnEnableFileOpenShortcut) Then
                Call ActivateFileOpenShortcut()
            Else
                Call DeactivateFileOpenShortcut()
            End If
        End If
    End Property
    
    Private Sub ActivateFileOpenShortcut()
        Call AssignFileOpenShortcut()
    End Sub
    
    Private Sub DeactivateFileOpenShortcut()
        Call ResetFileOpenShortcut()
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


' Region "SaveAsPDF" (Tastenkürzel "STRG+UMSCHALT+p" startet den Dialog "Als PDF veröffentlichen")
    
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

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:
