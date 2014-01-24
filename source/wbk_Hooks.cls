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
'---------------------------------------------------------------------------------------------------
 
Option Explicit

Private SyncWorkDir                             As SyncWorkDir
Private blnEnableSyncWorkDir                    As Boolean
Private blnEnableSyncWorkDirInitialized         As Boolean

Private blnEnableFileNewDirect                  As Boolean
Private blnEnableFileNewDirectInitialized       As Boolean

Private blnEnableConditionalFormat              As Boolean
Private blnEnableConditionalFormatInitialized   As Boolean


Private strVBAHostNameVersion   As String
Private strAddInName            As String


' Initialisierung
Private Sub Workbook_Open()
    ' Initialisierung der Klassen-Hooks.
    'If (Me.EnableConditionalFormat) Then Call ActivateConditionalFormat
    'If (Me.EnableFileNewDirect) Then Call ActivateFileNewDirect
    If (Me.EnableSyncWorkDir) Then Call ActivateSyncWorkDir
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
        getRibbon().Invalidate
    End Sub
    
    Private Sub DeactivateConditionalFormat()
        getRibbon().Invalidate
    End Sub
'


' Region "FileNewDirect" ("Datei -> Neu..." führt direkt zum Vorlagendialog)
    
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

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
