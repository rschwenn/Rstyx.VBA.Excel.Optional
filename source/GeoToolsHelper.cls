VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoToolsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===============================================================================
'  GeoToolsHelper.cls
'===============================================================================
' 
' Helps starting and stopping the GeoTools-Add-In.
'
'-------------------------------------------------------------------------------
                                                                                                                                      
Option Explicit

Const   AddinFileName       As String = "GeoTools.xlam"
Const   ExtraSearchDir      As String = ""

Const   DirTypeInstall      As String = "Installations-Verzeichnis"
Const   DirTypeLibrary      As String = "AddIn-Verzeichnis"
Const   DirTypeStartup      As String = "Startup-Verzeichnis"
Const   DirTypeExtra        As String = "Vereinbartes Verzeichnis"
Const   DirTypeThisAddIn    As String = "Verzeichnis des Optional Add-Ins"


Private blnIsAddInFound     As Boolean
Private blnIsAddInInstalled As Boolean
Private blnIsAddInOpen      As Boolean
Private blnIsAddInListed    As Boolean
Private strAddInFullName    As String
Private strAddInDirName     As String
Private strAddInDirType     As String
Private strAddInError       As String
Private strActionMessage    As String
Private oAddin              As AddIn


' Region "Properties"
    
    Property Get IsAddInFound() As Boolean
        IsAddInFound = blnIsAddInFound
    End Property
    
    Property Get IsAddInInstalled() As Boolean
        IsAddInInstalled = blnIsAddInInstalled
    End Property
    
    Property Get IsAddInOpen() As Boolean
        IsAddInOpen = blnIsAddInOpen
    End Property
    
    Property Get IsAddInListed() As Boolean
        ' True, if Add-In is listed in the GUI Addin list (resp. Application.Addins).
        IsAddInListed = blnIsAddInListed
    End Property
    
    Property Get IsAddInAutoLoad() As Boolean
        ' True, if Add-In would be startet automatically at next Excel start.
        IsAddInAutoLoad = (Me.IsAddInInstalled Or (Me.AddInDirType = DirTypeStartup))
    End Property
    
    Property Get AddInName() As String
        AddInName = AddinFileName
    End Property
    
    Property Get AddInFullName() As String
        AddInFullName = strAddInFullName
    End Property
    
    Property Get AddInDirName() As String
        AddInDirName = strAddInDirName
    End Property
    
    Property Get AddInDirType() As String
        AddInDirType = strAddInDirType
    End Property
    
    Property Get AddInError() As String
        AddInError = strAddInError
    End Property
    
    Property Get ActionMessage() As String
        ActionMessage = strActionMessage
    End Property
'

' Region "Public Methods"
    
    Public Sub StartAddIn()
        ' Tries to start the AddIn.
        On Error Goto Catch
        
        If (Me.IsAddInOpen) Then
            
            Call SetActionMessage("Add-In ist bereits aktiv => nichts zu tun.")
            
        ElseIf (Me.IsAddInInstalled) Then
            ' In fact the Add-In should has been loaded already.
            oAddin.Installed = False
            oAddin.Installed = True
            Call SetActionMessage("Add-In wurde neu geladen (mit Excel's Add-In-Mechanismus).")
            
        ElseIf (Me.AddInDirType = DirTypeInstall) Then
            ' Add-In is available in Add-In list.
            oAddin.Installed = True
            Call SetActionMessage("Add-In wurde geladen (mit Excel's Add-In-Mechanismus).")
            
        ElseIf (Me.AddInDirType = DirTypeStartup) Then
            ' Add-In isn't available in Add-In list, but has been found in a xlstart folder.
            ' (Since it isn't loaded, it should has been unloaded manually before).
            Dim AddinWbk As Workbook
            Set AddinWbk = ThisWorkbook.Application.Workbooks.Open(FileName:=Me.AddInFullName, ReadOnly:=True, UpdateLinks:=0 , AddToMru:=False)
            Call SetActionMessage("Add-In wurde geladen (durch Öffnen der Datei).")
            
        ElseIf (Me.IsAddInFound) Then
            ' Add-In isn't available in Add-In list, but has been found in a non-startup folder.
            Set oAddin = ThisWorkbook.Application.AddIns.Add(Filename:=Me.AddInFullName, CopyFile:=False)
            oAddin.Installed = True
            Call SetActionMessage("Add-In wurde installiert und geladen (mit Excel's Add-In-Mechanismus).")
            
        Else
            ' Add-In isn't available in Add-In list and hasn't been found.
            Call SetActionMessage("Programmfehler: Add-In kann nicht gestartet werden, da sein Speicherort unbekannt ist.")
        End If
        
        Call GetAddinStatus()
        Exit Sub
        
        Catch:
        Call SetActionMessage(Err.Description)
        Err.Clear
        Call GetAddinStatus()
    End Sub
    
    Public Sub StopAddIn()
        ' Tries to stop the AddIn.
        On Error Goto Catch
        
        If (Not Me.IsAddInOpen) Then
            Call SetActionMessage("Add-In ist nicht aktiv => nichts zu tun.")
            
        ElseIf (Not Me.IsAddInInstalled) Then
            ' Not installed but loaded Add-In (via startup folder or manually) => Close workbook.
            Dim AddinWbk As Workbook
            Set AddinWbk = ThisWorkbook.Application.Workbooks(AddinFileName)
            AddinWbk.Close False
            Call SetActionMessage("Add-In wurde entladen (durch Schließen der Datei).")
        Else
            ' Normal Add-In: installed and loaded.
            oAddin.Installed = False
            Call SetActionMessage("Add-In wurde entladen und deaktiviert (mit Excel's Add-In-Mechanismus).")
        End If
        
        Call GetAddinStatus()
        Exit Sub
        
        Catch:
        Call SetActionMessage(Err.Description)
        Err.Clear
        Call GetAddinStatus()
    End Sub
    
    Public Sub InvalidateStatus()
        Call GetAddinStatus()
    End Sub
    
    Public Sub DebugAddins()
        Dim oCurrentAddin    As AddIn
        
        For Each oCurrentAddin In ThisWorkbook.Application.AddIns2
            Debug.Print "Add-In:  Name=" & oCurrentAddin.Name      & vbTab & _
                        "Installed="     & oCurrentAddin.Installed & vbTab & _
                        "IsOpen="        & oCurrentAddin.IsOpen    & vbTab & _
                        "FullName="      & oCurrentAddin.FullName
        Next
        
        Debug.Print "Application.LibraryPath = "     & vbTab & ThisWorkbook.Application.LibraryPath
        Debug.Print "Application.UserLibraryPath = " & vbTab & ThisWorkbook.Application.UserLibraryPath
    End Sub
'

' Region "Private Methods"
    
    Private Sub GetAddinStatus()
        ' Gets the AddIn's status and sets status properties.
        
        Debug.Print "GeoToolsHelper.GetAddinStatus() startet."
        
        blnIsAddInFound     = False
        blnIsAddInInstalled = False
        blnIsAddInOpen      = False
        blnIsAddInListed    = False
        strAddInFullName    = "?"
        strAddInDirName     = "?"
        strAddInDirType     = "?"
        strAddInError       = ""
        
        Set oAddin = GetAddin()
        
        If (Not oAddin Is Nothing) Then
            
            Debug.Print "GeoToolsHelper.GetAddinStatus(): " & Me.AddInName & " in Add-In-Liste gefunden."
            
            blnIsAddInFound     = (Not (Dir(oAddin.FullName) = ""))
            blnIsAddInListed    = True
            blnIsAddInInstalled = oAddin.Installed
            blnIsAddInOpen      = oAddin.IsOpen
            'strAddInName       = oAddin.Name
            strAddInFullName    = oAddin.FullName
            strAddInDirName     = oAddin.Path
            
            If (Not Me.IsAddInFound) Then
                strAddInError = "Add-In ist mit ungültigem Pfad registriert => Bitte manuell bereinigen."
            End If
            
            If (oAddin.Installed Or (Not oAddin.IsOpen)) Then
                strAddInDirType = DirTypeInstall
            Else
                ' Determine directory type for non-installed but loaded Add-In.
                Select Case Lcase(oAddin.Path)
                    Case Lcase(ThisWorkbook.Application.LibraryPath)     :  strAddInDirType = DirTypeLibrary
                    Case Lcase(ThisWorkbook.Application.UserLibraryPath) :  strAddInDirType = DirTypeLibrary
                    Case Lcase(ThisWorkbook.Application.StartupPath)     :  strAddInDirType = DirTypeStartup
                    Case Lcase(ThisWorkbook.Application.AltStartupPath)  :  strAddInDirType = DirTypeStartup
                    Case Lcase(ThisWorkbook.Path)                        :  strAddInDirType = DirTypeThisAddIn
                    Case Lcase(ExtraSearchDir)                           :  strAddInDirType = DirTypeExtra
                End Select
            End If
        Else
            ' Add-In isn't listed or otherwise loaded.
            ' => Search for Add-In file in several directories.
            Debug.Print "GeoToolsHelper.GetAddinStatus(): " & Me.AddInName & " nicht in Add-In-Liste gefunden oder anderweitig geladen."
            
            ' Library directories.
            blnIsAddInFound = LookForAddin(ThisWorkbook.Application.LibraryPath, DirTypeLibrary)
            If (Not blnIsAddInFound) Then blnIsAddInFound = LookForAddin(ThisWorkbook.Application.UserLibraryPath, DirTypeLibrary)
            
            ' Startup directories.
            If (Not blnIsAddInFound) Then blnIsAddInFound = LookForAddin(ThisWorkbook.Application.StartupPath, DirTypeStartup)
            If (Not blnIsAddInFound) Then blnIsAddInFound = LookForAddin(ThisWorkbook.Application.AltStartupPath, DirTypeStartup)
            
            ' Directory of this AddIn.
            If (Not blnIsAddInFound) Then blnIsAddInFound = LookForAddin(ThisWorkbook.Path, DirTypeThisAddIn)
            
            ' Extra directory.
            If (Not blnIsAddInFound) Then blnIsAddInFound = LookForAddin(ExtraSearchDir, DirTypeExtra)
        End If
    End Sub
    
    Private Function LookForAddin(ByVal Directory As String, ByVal DirType As String) As Boolean
        ' --------------------------------------------------------------------------
        ' Looks for Me.AddInName in given directory.
        ' 
        ' Input:  Directory: Folder to look into.
        '         DirType:   If file is found, then Me.AddInDirType will be set to this.
        ' 
        ' Side effects:      If file is found, then Me.AddInFullName will be set to file path.
        ' 
        ' Returns:           True, if file has been found.
        ' --------------------------------------------------------------------------
        Dim FileFound   As Boolean
        Dim FullName    As String
        
        If (Not (Directory = "")) Then
            
            If (Right(Directory, 1) = ThisWorkbook.Application.PathSeparator) Then
                Directory = Left(Directory, Len(Directory)- 1)
            End If
            
            FullName = Directory & ThisWorkbook.Application.PathSeparator & AddinFileName
            
            If (Not (Dir(FullName) = "")) Then
                FileFound        = True
                strAddInFullName = FullName
                strAddInDirName  = Directory
                strAddInDirType  = DirType
            End If
        End If
        
        LookForAddin = FileFound
    End Function
    
    Private Function GetAddin() As AddIn
        ' Returns the Add-In, even if the file doesn't exists (or Nothing).
        Dim oCurrentAddin   As AddIn
        Dim oAddin          As AddIn
        
        Debug.Print "Searching for available and otherwise loaded AddIns (Application.AddIns2) ..."
        For Each oCurrentAddin In ThisWorkbook.Application.AddIns2
            
            Debug.Print "Name"      & "=" & oCurrentAddin.Name      & vbTab & _
                        "Installed" & "=" & oCurrentAddin.Installed & vbTab & _
                        "IsOpen"    & "=" & oCurrentAddin.IsOpen    & vbTab & _
                        "FullName"  & "=" & oCurrentAddin.FullName
            '
            If (Lcase(Right(oCurrentAddin.FullName, Len(AddinFileName))) = Lcase(AddinFileName)) Then
                Debug.Print "=> That's the target AddIn."
                'If (Dir(oCurrentAddin.FullName) = "") Then
                '    Debug.Print "=> Target AddIn doesn't exists."
                'Else
                '    Debug.Print "=> Target AddIn exists."
                    Set oAddin = oCurrentAddin
                    Exit For
                'End If
            End If
        Next
        
        'If (oAddin Is Nothing) Then
            'Debug.Print "Searching for otherwise opened AddIns (Application.AddIns2) ..."
            'For Each oCurrentAddin In ThisWorkbook.Application.AddIns2
            '    
            '    Debug.Print "Name"      & "=" & oCurrentAddin.Name      & vbTab & _
            '                "Installed" & "=" & oCurrentAddin.Installed & vbTab & _
            '                "IsOpen"    & "=" & oCurrentAddin.IsOpen    & vbTab & _
            '                "FullName"  & "=" & oCurrentAddin.FullName
            '    '
            '    If (Instr(lcase(oCurrentAddin.FullName), lcase(AddinFileName)) > 0) Then
            '        Debug.Print "=> That's the target AddIn."
            '        Debug.Print "=> Add the AddIn to the available Add-Ins list."
            '        Call ThisWorkbook.Application.AddIns.Add(Filename:=oCurrentAddin.FullName)  *** FEHLER !?!?
            '        Set oAddin = oCurrentAddin
            '        Exit For
            '    End If
            'Next
        'End If
        
        Set GetAddin = oAddin
    End Function
    
    Private Sub SetActionMessage(Message As String)
        ' Sets Me.ActionMessage.
        On Error GoTo Catch
        strActionMessage = Message
        
        Exit Sub
        Catch:
        Err.Clear
    End Sub
    
    Private Sub ClearActionMessage()
        ' Clears Me.ActionMessage.
        On Error GoTo Catch
        strActionMessage = ""
        
        Exit Sub
        Catch:
        Err.Clear
    End Sub
'

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
