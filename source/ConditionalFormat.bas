Attribute VB_Name = "ConditionalFormat"
'---------------------------------------------------------------------------------------------------
' ConditionalFormat.bas (Robert Schwenn)
' 
' Der Kontextmenüeintrag "Bedingte Formatierung" wird via XML angelegt.
' Dessen Sichtbarkeit wird gesteuert via Callback "getVisibleConditionalFormat()"
' das die Eigenschaft "EnableConditionalFormat" zurückgibt.
'---------------------------------------------------------------------------------------------------

Option Explicit

' Ribbon-Callback
Public Sub getVisibleConditionalFormat(control As IRibbonControl, ByRef visible)
    
    visible = ThisWorkbook.EnableConditionalFormat
End Sub


' for jEdit:  :collapseFolds=1::tabSize=4::indentSize=4:
