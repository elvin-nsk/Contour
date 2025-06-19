VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public IsOk As Boolean
Public IsCancel As Boolean

Public OutlineColor As Color
Public FillColor As Color

Public OffsetHandler As TextBoxHandler
Public OffsetHandler2 As TextBoxHandler
Public OutlineWidthHandler As TextBoxHandler
Public NameHandler As TextBoxHandler

'===============================================================================

Private Sub UserForm_Initialize()
    Caption = LocalizedStrings("MainView.Caption") & " v" & APP_VERSION
    Logo.ControlTipText = APP_URL
    Localize
    
    Set OutlineColor = CreateColor
    Set FillColor = CreateColor
    Set OffsetHandler = _
        TextBoxHandler.SetDouble(TextBoxOffset, -10000#, 10000#)
    Set OffsetHandler2 = _
        TextBoxHandler.SetDouble(TextBoxOffset2, -10000#, 10000#)
    Set OutlineWidthHandler = _
        TextBoxHandler.SetDouble(TextBoxOutlineWidth, 0.001)
    Set NameHandler = _
        TextBoxHandler.SetString( _
                           TextBoxName, 1, 16, _
                           LocalizedStrings("MainView.TextBoxName.Default") _
                       )
    
End Sub

Private Sub UserForm_Activate()
    ButtonOutlineColor.BackColor = ColorToRGB(OutlineColor)
    ButtonFillColor.BackColor = ColorToRGB(FillColor)
End Sub

Private Sub ButtonOutlineColor_Click()
    If PickColor(ButtonOutlineColor, OutlineColor) Then OptionMakeOutline = True
End Sub

Private Sub ButtonFillColor_Click()
    If PickColor(ButtonFillColor, FillColor) Then
        OptionMakeFill = True
        OptionFillColor = True
    End If
End Sub

Private Sub OptionResultAbove_Click()
    OptionResultBelow = Not OptionResultAbove
End Sub

Private Sub OptionResultBelow_Click()
    OptionResultAbove = Not OptionResultBelow
End Sub

Private Sub TextBoxOffset_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then FormОК
End Sub

Private Sub TextBoxOutlineWidth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then FormОК
End Sub

Private Sub TextBoxName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then FormОК
End Sub

Private Sub ButtonOk_Click()
    FormОК
End Sub

Private Sub ButtonCancel_Click()
    FormCancel
End Sub

Private Sub Logo_Click()
    With VBA.CreateObject("WScript.Shell")
        .Run APP_URL
    End With
End Sub

'===============================================================================

Private Sub Localize()

    FrameContour.Caption = LocalizedStrings("MainView.FrameContour")
    LabelOffset.Caption = LocalizedStrings("MainView.LabelOffset")
    LabelOffsetUnits.Caption = LocalizedStrings("MainView.LabelOffsetUnits")
    OptionMakeOutline.Caption = LocalizedStrings("MainView.OptionMakeOutline")
    LabelOutlineWidth.Caption = LocalizedStrings("MainView.LabelOutlineWidth")
    LabelOutlineUnits.Caption = LocalizedStrings("MainView.LabelOutlineUnits")
    OptionMakeFill.Caption = LocalizedStrings("MainView.OptionMakeFill")
    OptionMatchColor.Caption = LocalizedStrings("MainView.OptionMatchColor")
    OptionTrace.Caption = LocalizedStrings("MainView.OptionTrace")
    OptionRoundCorners.Caption = LocalizedStrings("MainView.OptionRoundCorners")
    
    OptionSecondaryContour.Caption = LocalizedStrings("MainView.OptionSecondaryContour")
    LabelOffset2.Caption = LocalizedStrings("MainView.LabelOffset2")
    LabelOffsetUnits2.Caption = LocalizedStrings("MainView.LabelOffsetUnits2")
    OptionRoundCorners2.Caption = LocalizedStrings("MainView.OptionRoundCorners2")
    
    FrameSource.Caption = LocalizedStrings("MainView.FrameSource")
    OptionSourceAsOne.Caption = LocalizedStrings("MainView.OptionSourceAsOne")
    OptionSourceAsIs.Caption = LocalizedStrings("MainView.OptionSourceAsIs")
    OptionSourceWithinGroups.Caption = LocalizedStrings("MainView.OptionSourceWithinGroups")
    
    FrameResult.Caption = LocalizedStrings("MainView.FrameResult")
    OptionResultAbove.Caption = LocalizedStrings("MainView.OptionResultAbove")
    OptionResultBelow.Caption = LocalizedStrings("MainView.OptionResultBelow")
    OptionResultAsObjects.Caption = LocalizedStrings("MainView.OptionResultAsObjects")
    OptionResultAsGroup.Caption = LocalizedStrings("MainView.OptionResultAsGroup")
    OptionResultAsLayer.Caption = LocalizedStrings("MainView.OptionResultAsLayer")
    LabelName.Caption = LocalizedStrings("MainView.LabelName")
    
    ButtonOk.Caption = LocalizedStrings("MainView.ButtonOk")

End Sub

Private Function ColorToRGB(ByVal Color As Color) As Long
    With CreateColor
        .CopyAssign Color
        .ConvertToRGB
        ColorToRGB = VBA.RGB(.RGBRed, .RGBGreen, .RGBBlue)
    End With
End Function

Private Function PickColor( _
                     ByVal TargetButton As MSForms.CommandButton, _
                     ByVal TargetColor As Color _
                 ) As Boolean
    Dim PickedColor As Color
    Set PickedColor = CreateColor
    PickedColor.CopyAssign TargetColor
    If Not PickedColor.UserAssignEx Then Exit Function
    TargetColor.CopyAssign PickedColor
    TargetButton.BackColor = ColorToRGB(PickedColor)
    PickColor = True
End Function

Private Sub FormОК()
    Me.Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Me.Hide
    IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Сancel = True
        FormCancel
    End If
End Sub
