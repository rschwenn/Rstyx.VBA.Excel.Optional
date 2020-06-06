Attribute VB_Name = "FullScreen"
'---------------------------------------------------------------------------------------------------
' FullScreen.bas (Robert Schwenn)
' 
' Das Makro "Optional.FullScreen.ToggleFullScreen" schalten den Vollbildmodus um.
' Diesem Makro kann mit den anderen beiden Makros ***FullScreenShortcut()
' das Tastenkürzel "F11" zugewiesen bzw. entzogen werden.
'---------------------------------------------------------------------------------------------------

Option Explicit


' Tastenkürzel "F11" wird dem Makro "ToggleFullScreen" zugewiesen. 
Sub AssignFullScreenShortcut()
  Application.OnKey "{F11}", "ToggleFullScreen"
End Sub

' Tastenkürzel "F11" wird auf Standard (Diagramm einfügen) zurückgesetzt. 
Sub ResetFullScreenShortcut()
  Application.OnKey "{F11}"
End Sub


' Umschalten des Vollbildmodus.
Sub ToggleFullScreen()
    'On Error Resume Next
    
    If (Not ThisWorkbook.FullScreenExtended) Then
        
        'Debug.Print "DisplayFullScreen = " & Application.DisplayFullScreen
        Application.DisplayFullScreen = (Not Application.DisplayFullScreen)
    Else
        'Debug.Print "WindowState = " & Application.WindowState
        
        If (Application.WindowState = xlMaximized) Then
            Application.WindowState = xlNormal
        Else
            Application.WindowState = xlMaximized
        End If
    End If
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
