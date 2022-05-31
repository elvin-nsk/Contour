VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
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
Public OutlineWidthHandler As TextBoxHandler
Public NameHandler As TextBoxHandler

'===============================================================================

Private Sub UserForm_Initialize()
    Caption = APP_NAME
    Logo.ControlTipText = APP_URL
    'Localize
    
    Set OutlineColor = CreateColor
    Set FillColor = CreateColor
    Set OffsetHandler = _
        TextBoxHandler.SetDouble(TextBoxOffset, 0.001)
    Set OutlineWidthHandler = _
        TextBoxHandler.SetDouble(TextBoxOutlineWidth, 0.001)
    Set NameHandler = _
        TextBoxHandler.SetString(TextBoxName, 1, 16)
    
End Sub

Private Sub UserForm_Activate()
    ButtonOutlineColor.BackColor = ColorToRGB(OutlineColor)
    ButtonFillColor.BackColor = ColorToRGB(FillColor)
End Sub

Private Sub ButtonOutlineColor_Click()
    PickColor ButtonOutlineColor, OutlineColor
End Sub

Private Sub ButtonFillColor_Click()
    PickColor ButtonFillColor, FillColor
End Sub

Private Sub OptionResultAbove_Click()
    OptionResultBelow = Not OptionResultAbove
End Sub

Private Sub OptionResultBelow_Click()
    OptionResultAbove = Not OptionResultBelow
End Sub

Private Sub TextBoxOffset_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Form问
End Sub

Private Sub TextBoxOutlineWidth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Form问
End Sub

Private Sub TextBoxName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then Form问
End Sub

Private Sub ButtonOk_Click()
    Form问
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
    OptionMakeFill.Caption = LocalizedStrings("MainView.OptionMakeFill")
    OptionMatchColor.Caption = LocalizedStrings("MainView.OptionMatchColor")
    OptionTrace.Caption = LocalizedStrings("MainView.OptionTrace")
    
    FrameSource.Caption = LocalizedStrings("MainView.FrameSource")
    OptionSourceAsOne.Caption = LocalizedStrings("MainView.OptionSourceAsOne")
    OptionSourceAsIs.Caption = LocalizedStrings("MainView.OptionSourceAsIs")
    OptionSourceWithinGroups.Caption = LocalizedStrings("MainView.OptionSourceWithinGroups")
    
    FrameResult.Caption = LocalizedStrings("MainView.FrameResult")
    OptionResultAsObjects.Caption = LocalizedStrings("MainView.OptionResultAsObjects")
    OptionResultAsGroup.Caption = LocalizedStrings("MainView.OptionResultAsGroup")
    OptionResultAsLayer.Caption = LocalizedStrings("MainView.OptionResultAsLayer")
    
    ButtonOk.Caption = LocalizedStrings("MainView.ButtonOk")
    ButtonCancel.Caption = LocalizedStrings("MainView.ButtonCancel")

End Sub

Private Function ColorToRGB(ByVal Color As Color) As Long
    With CreateColor
        .CopyAssign Color
        .ConvertToRGB
        ColorToRGB = VBA.RGB(.RGBRed, .RGBGreen, .RGBBlue)
    End With
End Function

Private Sub PickColor( _
                ByVal TargetButton As MSForms.CommandButton, _
                ByVal TargetColor As Color _
            )
    Dim PickedColor As Color
    Set PickedColor = CreateColor
    PickedColor.CopyAssign TargetColor
    If Not PickedColor.UserAssignEx Then Exit Sub
    TargetColor.CopyAssign PickedColor
    TargetButton.BackColor = ColorToRGB(PickedColor)
End Sub

Private Sub Form问()
    Me.Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Me.Hide
    IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(補ncel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        補ncel = True
        FormCancel
    End If
End Sub
