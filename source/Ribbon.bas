Attribute VB_Name = "Ribbon"
'===============================================================================
' Modul Ribbon                                                                  
'===============================================================================

Option Explicit



Private oRibbon As IRibbonUI


' Region "Ribbon-Objekt (Referenz, Update)"
    
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal cBytes As Long)
    
    ' Initialisierung der RibbonUI: Speichern einer Referenz auf das Ribbon-Objekt
    ' und als Backup eines entsprechenden Integer-Zeigers in die Add-In-interne Tabelle.
    Public Sub OnOptionalRibbonLoad(ribbon As IRibbonUI)
        Set oRibbon = ribbon
        tabOptional.Range("A1").Value = ObjPtr(ribbon)
    End Sub
    
    ' Beziehen einer Referenz auf das Ribbon-Objekt (Sollte auch nach Fehler im Add-In funktionieren).
    Function getOptionalRibbon() As IRibbonUI
        ' "oRibbon" ist normalerweise nur dann "Nothing", wenn das AddIn wegen eines Fehlers gestoppt wurde.
        ' Dann kann der vorher gespeicherte Zeiger verwendet werden.
        ' ABER: Wenn das AddIn nicht schreibgeschützt ist, kann der Zeiger auch veraltet sein.
        '       => Dann stürzt Excel ab und nichts geht mehr.
        If (oRibbon Is Nothing) Then
            If (ThisWorkbook.ReadOnly) Then
                Dim ribbonPointer As LongPtr
                ribbonPointer = tabOptional.Range("A1").value
                If (ribbonPointer > 0) Then
                    On Error Resume Next  ' Nützt nix!
                    Call CopyMemory(oRibbon, ribbonPointer, LenB(ribbonPointer))
                    On Error GoTo 0
                End If
            End If
        End If
        
        Set getOptionalRibbon = oRibbon
    End Function
    
    ' Status-Aktualisierung aller Ribbon-Steuerelemente erzwingen.
    Public Sub UpdateOptionalRibbon()
        On Error Resume Next
        getOptionalRibbon().Invalidate
        'call ClearStatusBarDelayed(3)
        On Error Goto 0
    End Sub
    
' End Region


' Region "Backstage"
    
    Sub BackstageOnShow(control As IRibbonControl)
        Debug.Print "Ribbon.BackstageOnShow()"
        Call ThisWorkbook.GeoToolsHelper.InvalidateStatus()
    End Sub
    
' End Region

' Region "Checkboxes"
    
    ' Response to a click on a checkbox.
    Sub OptionalCheckboxAction(control As IRibbonControl, pressed As Boolean)
        On Error Resume Next
        Select Case control.ID
            Case "EnableFileNewShortcutCheckbox"   :  ThisWorkbook.EnableFileNewShortcut   = pressed
            Case "EnableFileOpenShortcutCheckbox"  :  ThisWorkbook.EnableFileOpenShortcut  = pressed
            Case "EnableSaveAsPDFCheckbox"         :  ThisWorkbook.EnableSaveAsPDF         = pressed
            Case "EnableConditionalFormatCheckbox" :  ThisWorkbook.EnableConditionalFormat = pressed
            Case "EnableFileNewButtonCheckbox"     :  ThisWorkbook.EnableFileNewButton     = pressed
        End select
        Call UpdateOptionalRibbon
        On Error Goto 0
    End Sub
    
    ' Get status of a checkbox.
    Sub OptionalCheckboxGetPressed(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        Select Case control.ID
            Case "EnableFileNewShortcutCheckbox"   :  returnedVal = ThisWorkbook.EnableFileNewShortcut
            Case "EnableFileOpenShortcutCheckbox"  :  returnedVal = ThisWorkbook.EnableFileOpenShortcut
            Case "EnableSaveAsPDFCheckbox"         :  returnedVal = ThisWorkbook.EnableSaveAsPDF
            Case "EnableConditionalFormatCheckbox" :  returnedVal = ThisWorkbook.EnableConditionalFormat
            Case "EnableFileNewButtonCheckbox"     :  returnedVal = ThisWorkbook.EnableFileNewButton
        End select
        On Error Goto 0
    End Sub
    
' End Region


' Region "Labels"
    
    ' Get text for a Label.
    Sub OptionalLabelGetText(control As IRibbonControl, ByRef returnedVal)
        On Error Resume Next
        Debug.Print "Ribbon.OptionalLabelGetText():  control.ID = " & control.ID
        Select Case control.ID
            Case "GtInfoLabelFilename" :  returnedVal = ThisWorkbook.GeoToolsHelper.AddInName
            Case "GtInfoLabelDirName"  :  returnedVal = ThisWorkbook.GeoToolsHelper.AddInDirName
            Case "GtInfoLabelDirType"  :  returnedVal = ThisWorkbook.GeoToolsHelper.AddInDirType
            Case "GtInfoLabelAutoLoad" :  returnedVal = ThisWorkbook.GeoToolsHelper.IsAddInAutoLoad
            Case "GtInfoLabelLoaded"   :  returnedVal = ThisWorkbook.GeoToolsHelper.IsAddInOpen
            Case "GtInfoLabelAction"   :  returnedVal = ThisWorkbook.GeoToolsHelper.ActionResult
        End select
        On Error Goto 0
    End Sub
    
' End Region

' Region "Action Buttons"
    
    Sub OptionalButtonAction(ByVal control As IRibbonControl)
        'On Error Resume Next
        Debug.Print "Ribbon.OptionalButtonAction():  control.ID = " & control.ID
        Select Case control.ID
            Case "GtStartButton"   : Call ThisWorkbook.GeoToolsHelper.StartAddIn
            Case "GtStopButton"    : Call ThisWorkbook.GeoToolsHelper.StopAddIn
            Case Else              : Msgbox "Ribbon.OptionalButtonAction(): Unbekannte Control.ID = " & control.ID
        End select
        Call UpdateOptionalRibbon
        'On Error Goto 0
    End Sub
    
    Sub GetEnabledButtonStart(control As IRibbonControl, ByRef returnedVal)
        Debug.Print "Ribbon.GetEnabledButtonStart():  control.ID = " & control.ID
        returnedVal = (ThisWorkbook.GeoToolsHelper.IsAddInFound And (Not ThisWorkbook.GeoToolsHelper.IsAddInOpen))
    End Sub
    
    Sub GetEnabledButtonStop(control As IRibbonControl, ByRef returnedVal)
        Debug.Print "Ribbon.GetEnabledButtonStop():  control.ID = " & control.ID
        returnedVal = ThisWorkbook.GeoToolsHelper.IsAddInOpen
    End Sub
    
' End Region

' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
