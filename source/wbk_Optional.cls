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
' 24.11.2008 v1.0   Der Men�punkt "Datei -> Neu..." wird hier verarbeitet:
'                     Statt den Aufgabenbereich anzuzeigen, startet direkt der
'                     Dialog "Vorlage w�hlen" (wie Excel 97) im Listenmodus
' 04.01.2014 v2.0   - Umstrukturiert: 
'                     - Je Hook ein eigenes Modul
'                     - In wbk_Hooks.cls: je Hook eine Aktivierungs- und Deaktivierungsroutine
'                       sowie eine Eigenschaft (Hauptschalter)
'                   - Anpassung f�r Excel 2010:
'                     - Hook f�r "Datei -> Neu..." umgestrickt.
'                     - Synchronisierung des Excel-Arbeitsverzeichnisses mit dem Speicherort
'                       der aktiven Arbeitsmappe, damit der �ffnen-Dialog dort startet.
'                       (Der Speichern-unter-Dialog startet ohnehin dort).
'                     - Einstellungsdialog in Register-Tab "Spezial"
' 24.01.2014 v2.1   - Neu: Kontextmen�eintrag "Bedingte Formatierung"
' 16.02.2014 v2.2   - Ribbon-Kode: Excel-weit eindeutige Routinenamen (Problemvermeidung)
' 20.11.2016 v2.3   - Bugfix: Synchronisierung des Excel-Arbeitsverzeichnisses arbeitet jetzt
'                     auch mit UNC-Pfaden (bisher trat ein Fehler auf)
' 30.11.2016 v2.4   - Neu: PDF-Export per Tastenk�rzel "STRG+UMSCHALT+p".
' 11.05.2020 v3.0   - Umstellung auf Office 2016 (VBA 7).
'                   - Synchronisierung des Excel-Arbeitsverzeichnisses aufgegeben.
'                     (scheint nun Excel-Standard zu sein wegen SDI)
'                   - FileNewDirect ersetzt durch FileNewShortcut und FileNewButton.
'                     (Die Standardaktion "Datei -> Neu..." wird nicht mehr abgefangen)
'                   - Neu: FileNewOpenShortcut
' 21.05.2020 v3.1   - Add-In umbenannt von "Hooks" nach "Optional"
'                   - Oberfl�che jetzt im Backoffice als Tab => Dialog entfernt.
'                   - Synchronisierung des Excel-Arbeitsverzeichnisses entfernt.
' 04.06.2020 v3.2   - Unterst�tzung f�r GeoTools Add-In: Status-Anzeige und Start/Stop.
' 06.06.2020 v3.3   - Neu: Umschalten des Vollbildmodus.
' 07.12.2020 v3.4   - FileOpenDirect: Dialog startet im Verzeichnis der aktiven Arbeitsmappe
' 07.02.2021 v3.5   - Neu: Klassischer "Datei Speichern als"-Dialog per Tastenk�rzel "STRG+UMSCHALT+s"
' 17.02.2021 v3.5.1 - GeoTools-Start-Knopf zeigt ScreenTip mit Hinweis, welche Ordner
'                     nach den GeoTools durhsucht werden.
'---------------------------------------------------------------------------------------------------
 
Option Explicit

Private blnEnableFileNewShortcut                As Boolean
Private blnEnableFileNewShortcutInitialized     As Boolean

Private blnEnableFileNewButton                  As Boolean
Private blnEnableFileNewButtonInitialized       As Boolean

Private blnEnableFileOpenShortcut               As Boolean
Private blnEnableFileOpenShortcutInitialized    As Boolean

Private blnEnableFileSaveAsShortcut             As Boolean
Private blnEnableFileSaveAsShortcutInitialized  As Boolean

Private blnEnableConditionalFormat              As Boolean
Private blnEnableConditionalFormatInitialized   As Boolean

Private blnEnableSaveAsPDF                      As Boolean
Private blnEnableSaveAsPDFInitialized           As Boolean

Private blnEnableFullScreenShortcut             As Boolean
Private blnEnableFullScreenShortcutInitialized  As Boolean

Private blnFullScreenExtended                   As Boolean
Private blnFullScreenExtendedInitialized        As Boolean

Private strVBAHostNameVersion                   As String
Private strAddInName                            As String

Private oGeoToolsHelper                         As GeoToolsHelper


' Initialisierung
Private Sub Workbook_Open()
    ' Initialisierung der Klassen.
    If (Me.EnableConditionalFormat)  Then Call ActivateConditionalFormat
    If (Me.EnableFileNewButton)      Then Call ActivateFileNewButton
    If (Me.EnableFileNewShortcut)    Then Call ActivateFileNewShortcut
    If (Me.EnableFileOpenShortcut)   Then Call ActivateFileOpenShortcut
    If (Me.EnableFileSaveAsShortcut) Then Call ActivateFileSaveAsShortcut
    If (Me.EnableFullScreenShortcut) Then Call ActivateFullScreenShortcut
    If (Me.EnableSaveAsPDF)          Then Call ActivateSaveAsPDF
    
    'Set oGeoToolsHelper = New GeoToolsHelper
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


' Region "FileOpenDialog" (klassischer �ffnen-Dialog)
    
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


' Region "FileSaveAsDialog" (klassischer Speichern als-Dialog)
    
    Public Property Get EnableFileSaveAsShortcut() As Boolean
        If (Not blnEnableFileSaveAsShortcutInitialized) Then
            On Error Resume Next
            blnEnableFileSaveAsShortcut = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableFileSaveAsShortcut", EnableFileSaveAsShortcutDefault))
            On Error Goto 0
            blnEnableFileSaveAsShortcutInitialized = True
        End If
        EnableFileSaveAsShortcut = blnEnableFileSaveAsShortcut
    End Property
    
    Public Property Let EnableFileSaveAsShortcut(NewValue As Boolean)
        If (blnEnableFileSaveAsShortcut Xor NewValue) Then
            blnEnableFileSaveAsShortcut = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableFileSaveAsShortcut", CStr(blnEnableFileSaveAsShortcut)
            On Error Goto 0
            If (blnEnableFileSaveAsShortcut) Then
                Call ActivateFileSaveAsShortcut()
            Else
                Call DeactivateFileSaveAsShortcut()
            End If
        End If
    End Property
    
    Private Sub ActivateFileSaveAsShortcut()
        Call AssignFileSaveAsShortcut()
    End Sub
    
    Private Sub DeactivateFileSaveAsShortcut()
        Call ResetFileSaveAsShortcut()
    End Sub
'


' Region "ToggleFullScreen" (Umschalten Vollbildmodus)
    
    Public Property Get EnableFullScreenShortcut() As Boolean
        If (Not blnEnableFullScreenShortcutInitialized) Then
            On Error Resume Next
            blnEnableFullScreenShortcut = CBool(GetSetting(VBAHostNameVersion, AddInName, "EnableFullScreenShortcut", EnableFullScreenShortcutDefault))
            On Error Goto 0
            blnEnableFullScreenShortcutInitialized = True
        End If
        EnableFullScreenShortcut = blnEnableFullScreenShortcut
    End Property
    
    Public Property Let EnableFullScreenShortcut(NewValue As Boolean)
        If (blnEnableFullScreenShortcut Xor NewValue) Then
            blnEnableFullScreenShortcut = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "EnableFullScreenShortcut", CStr(blnEnableFullScreenShortcut)
            On Error Goto 0
            If (blnEnableFullScreenShortcut) Then
                Call ActivateFullScreenShortcut()
            Else
                Call DeactivateFullScreenShortcut()
            End If
        End If
    End Property
    
    
    Private Sub ActivateFullScreenShortcut()
        Call AssignFullScreenShortcut()
    End Sub
    
    Private Sub DeactivateFullScreenShortcut()
        Call ResetFullScreenShortcut()
    End Sub
    
    
    ' Option for ToggleFullScreen.
    Public Property Get FullScreenExtended() As Boolean
        If (Not blnFullScreenExtendedInitialized) Then
            On Error Resume Next
            blnFullScreenExtended = CBool(GetSetting(VBAHostNameVersion, AddInName, "FullScreenExtended", FullScreenExtendedDefault))
            On Error Goto 0
            blnFullScreenExtendedInitialized = True
        End If
        FullScreenExtended = blnFullScreenExtended
    End Property
    
    Public Property Let FullScreenExtended(NewValue As Boolean)
        If (blnFullScreenExtended Xor NewValue) Then
            blnFullScreenExtended = NewValue
            On Error Resume Next
            SaveSetting VBAHostNameVersion, AddInName, "FullScreenExtended", CStr(blnFullScreenExtended)
            On Error Goto 0
        End If
    End Property
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


' Region "GeoToolsHelper"
    
    Public Property Get GeoToolsHelper() As GeoToolsHelper
        If (oGeoToolsHelper Is Nothing) Then
            Set oGeoToolsHelper = New GeoToolsHelper
        End If
        Set GeoToolsHelper = oGeoToolsHelper
    End Function
'

' for jEdit:  :collapseFolds=2::tabSize=4::indentSize=4:
