VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrEnglishUS (1033)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
    
        .Add "Common.ErrNoDocument", "No active document"
        .Add "Common.ErrNoSelection", "Select shapes"
        .Add "Common.ErrDisabledLayer", _
            "Make active layer editable"
        .Add "Common.ErrConfig", "Configuration error"
        
        .Add "ProgressBar.TraceCaption", "Tracing..."
        .Add "ProgressBar.TraceMiddle", "of"
        
        .Add "MainView.FrameContour", "Contour"
        .Add "MainView.LabelOffset", "Offset"
        .Add "MainView.LabelOffsetUnits", "mm"
        .Add "MainView.OptionMakeOutline", "Outline"
        .Add "MainView.LabelOutlineWidth", "Outline width"
        .Add "MainView.LabelOutlineUnits", "mm"
        .Add "MainView.OptionMakeFill", "Fill"
        .Add "MainView.OptionMatchColor", "Match color"
        .Add "MainView.OptionTrace", "Trace bitmaps"
        
        .Add "MainView.FrameSource", "Process selection"
        .Add "MainView.OptionSourceAsOne", "as one shape"
        .Add "MainView.OptionSourceAsIs", "as separate shapes"
        .Add "MainView.OptionSourceWithinGroups", _
            "as separate shapes, count shapes in groups as separate"
        
        .Add "MainView.FrameResult", "Result is"
        .Add "MainView.OptionResultAbove", "above"
        .Add "MainView.OptionResultBelow", "below"
        .Add "MainView.OptionResultAsObjects", "every shape"
        .Add "MainView.OptionResultAsGroup", "shapes as group"
        .Add "MainView.OptionResultAsLayer", "current layer on separate layer"
        .Add "MainView.LabelName", "Name"
        .Add "MainView.TextBoxName.Default", "contour"
        
        .Add "MainView.ButtonOk", "Start"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
