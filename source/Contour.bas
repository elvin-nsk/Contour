Attribute VB_Name = "Contour"
Option Explicit
Private Const RELEASE As Boolean = True

'===============================================================================

'Отступ контура
Const ContourOffset As Double = 6 'мм

'Название слоя для контура
Const ContourName As String = "контур"

'===============================================================================

Private Type typeParams
  Offset As Double
  Name As String
End Type

Sub Start()
  
  If RELEASE Then On Error GoTo Catch
  
  If ActiveSelectionRange.Count = 0 Then Exit Sub
  
  lib_elvin.BoostStart "Установка контуров", RELEASE
  
  Dim Params As typeParams
  With Params
    .Offset = ContourOffset
    .Name = ContourName
  End With
  MainLoop ActiveSelectionRange, Params

Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  MsgBox "Ошибка: " & Err.Description, vbCritical
  Resume Finally

End Sub

Private Sub MainLoop(ShapeRange As ShapeRange, Params As typeParams)

  Dim Shape As Shape
  
  For Each Shape In ShapeRange
    If Shape.Type = cdrBitmapShape Then DoRoutine Shape, Params
  Next Shape

End Sub

Private Sub DoRoutine(Shape As Shape, Params As typeParams)

  Dim Trace As Shape
  Dim Cont As Shape
  
  Set Trace = DoTrace(Shape, Params)
  Set Cont = DoContour(Trace, Params)
  Trace.Delete
  lib_elvin.MoveToLayer Cont, ContourLayer(Params)
  Cont.Fill.ApplyNoFill
  Cont.Outline.Color.CMYKAssign 0, 0, 0, 100

End Sub

Private Function DoTrace(BitmapShape As Shape, Params As typeParams) As Shape
  With BitmapShape.Bitmap.Trace(cdrTraceLowQualityImage)
    .BackgroundRemovalMode = cdrTraceBackgroundAutomatic
    '.CornerSmoothness = 50
    '.DetailLevel = 10
    .SetColorCount 3
    .MergeAdjacentObjects = True
    .RemoveBackground = True
    .RemoveEntireBackColor = False
    Set DoTrace = .Finish.Group
  End With
End Function

Private Function DoContour(Shape As Shape, Params As typeParams) As Shape
  With Shape.CreateContour(cdrContourOutside)
    With .Contour
      .CornerType = cdrContourCornerRound
      .Offset = Params.Offset
      .Steps = 1
    End With
    Set DoContour = .Separate.FirstShape
  End With
End Function

Private Function ContourLayer(Params As typeParams) As Layer
  Set ContourLayer = ActivePage.Layers.Find(Params.Name)
  If ContourLayer Is Nothing Then
    Dim ALayer As Layer
    Set ALayer = ActiveLayer
    Set ContourLayer = ActivePage.CreateLayer(Params.Name)
    ContourLayer.MoveAbove ALayer
  End If
End Function
