VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Singleton

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const ReduceBitmapSize As Long = 1000
Private Const MmInInch = 25.4

'===============================================================================

Public Function ExcludeAndTraceBitmaps( _
                    ByVal Shapes As ShapeRange, _
                    ByVal OrderAbove As Boolean _
                ) As ShapeRange
    Set ExcludeAndTraceBitmaps = CreateShapeRange
    Dim BitmapShapes As ShapeRange
    Set BitmapShapes = Shapes.Shapes.FindShapes(Type:=cdrBitmapShape)
    If BitmapShapes.Count > 0 Then
        Shapes.RemoveRange BitmapShapes
        ExcludeAndTraceBitmaps.AddRange TraceBitmaps(BitmapShapes, OrderAbove)
    End If
End Function

Public Function TraceBitmaps( _
                    ByVal BitmapShapes As ShapeRange, _
                    ByVal OrderAbove As Boolean _
                ) As ShapeRange
    Set TraceBitmaps = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In BitmapShapes
        TraceBitmaps.Add TraceBitmap(Shape, OrderAbove)
    Next Shape
End Function

Public Function TraceBitmap( _
                    ByVal BitmapShape As Shape, _
                    ByVal OrderAbove As Boolean _
                ) As Shape

    Dim ShapeToProcess As Shape
    Dim DeleteShape As Boolean
                
    If BitmapShape.Bitmap.SizeWidth > ReduceBitmapSize _
    Or BitmapShape.Bitmap.SizeHeight > ReduceBitmapSize Then
        Set ShapeToProcess = BitmapShape.Duplicate
        DeleteShape = True
        Set ShapeToProcess = ShapeToProcess.ConvertToBitmapEx( _
            BitmapShape.Bitmap.Mode, , _
            BitmapShape.Bitmap.Transparent, _
            ReduceBitmapSize / (lib_elvin.GreaterDim(BitmapShape) / MmInInch) _
        )
    Else
        Set ShapeToProcess = BitmapShape
    End If
    
    Set TraceBitmap = Trace(ShapeToProcess)
    If OrderAbove Then
        TraceBitmap.OrderFrontOf BitmapShape
    Else
        TraceBitmap.OrderBackOf BitmapShape
    End If

    If DeleteShape Then ShapeToProcess.Delete
    
End Function

Public Function TryMakeBaseShape(ByVal Shape As Shape) As Shape
    On Error GoTo Catch
    Set TryMakeBaseShape = Shape.Layer.CreateCurve(Shape.Curve)
    Exit Function
Catch:
    Set TryMakeBaseShape = Shape.Layer.CreateRectangleRect(Shape.BoundingBox)
End Function

Private Function HasCurve(ByVal Shape As Shape) As Boolean
    On Error Resume Next
    If Not Shape.Curve Is Nothing Then HasCurve = True
    On Error GoTo 0
End Function

Public Function SeparateInvalidForContour(ByVal Shapes As ShapeRange) As ShapeRange
    Set SeparateInvalidForContour = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If Not ValidForContour(Shape) Then SeparateInvalidForContour.Add Shape
    Next Shape
    Shapes.RemoveRange SeparateInvalidForContour
End Function

Public Function ValidForContour(ByVal Shape As Shape) As Boolean
    Select Case Shape.Type
        Case cdrNoShape
        Case cdrRectangleShape
            ValidForContour = True
        Case cdrEllipseShape
            ValidForContour = True
        Case cdrCurveShape
            ValidForContour = True
        Case cdrPolygonShape
            ValidForContour = True
        Case cdrBitmapShape
            ValidForContour = True
        Case cdrTextShape
            ValidForContour = True
        Case cdrGroupShape
            ValidForContour = True
        Case cdrSelectionShape
        Case cdrGuidelineShape
        Case cdrBlendGroupShape
        Case cdrExtrudeGroupShape
        Case cdrOLEObjectShape
            ValidForContour = True
        Case cdrContourGroupShape
        Case cdrLinearDimensionShape
        Case cdrBevelGroupShape
            ValidForContour = True
        Case cdrDropShadowGroupShape
        Case cdr3DObjectShape
        Case cdrArtisticMediaGroupShape
        Case cdrConnectorShape
        Case cdrMeshFillShape
            ValidForContour = True
        Case cdrCustomShape
        Case cdrCustomEffectGroupShape
        Case cdrSymbolShape
        Case cdrHTMLFormObjectShape
        Case cdrHTMLActiveObjectShape
        Case cdrPerfectShape
        Case cdrEPSShape
    End Select
End Function

Public Function Contour( _
                    ByVal Shape As Shape, _
                    ByVal Offset As Double _
                 ) As Shape
    Set Contour = _
        Shape.CreateContour( _
                  Direction:=cdrContourOutside, _
                  Offset:=Offset, _
                  Steps:=1, _
                  CornerType:=cdrContourCornerRound _
              ).Separate.FirstShape
End Function

Public Sub AverageFill( _
               ByVal SourceShape As Shape, _
               ByVal TargetShape As Shape _
           )
    If SourceShape.Type = cdrGroupShape Then
        TargetShape.Fill.ApplyUniformFill _
            lib_elvin.GetAverageColorFromShapesFill(SourceShape.Shapes.All)
    Else
        TargetShape.Fill.CopyAssign SourceShape.Fill
    End If
End Sub

Public Function GetContourLayer( _
                    ByVal Name As String, _
                    ByVal OrderAbove As Boolean _
                ) As Layer
    Set GetContourLayer = ActivePage.Layers.Find(Name)
    If Not GetContourLayer Is Nothing Then Exit Function
    Dim ALayer As Layer
    Set ALayer = ActiveLayer
    Set GetContourLayer = ActivePage.CreateLayer(Name)
    If OrderAbove Then
        GetContourLayer.MoveAbove ALayer
    Else
        GetContourLayer.MoveBelow ALayer
    End If
End Function

Public Function CreateShapeRangeFromShape(ByVal Shape As Shape) As ShapeRange
    Set CreateShapeRangeFromShape = _
        Shape.Layer.Page.Parent.Parent.CreateShapeRangeFromArray(Shape)
End Function

Public Sub ThrowConfigError()
    VBA.Err.Raise CustomError, , "������ � ����������"
End Sub

'===============================================================================

Private Function Trace(ByVal BitmapShape As Shape) As Shape
    With BitmapShape.Bitmap.Trace(cdrTraceLowQualityImage)
        .BackgroundRemovalMode = cdrTraceBackgroundAutomatic
        '.CornerSmoothness = 50
        .DetailLevelPercent = 10
        '.SetColorCount 8
        .MergeAdjacentObjects = True
        .RemoveBackground = True
        .RemoveEntireBackColor = False
        Set Trace = .Finish.Group
    End With
End Function

Private Sub Class_Initialize()
    If Not Me Is Common Then Err.Raise 425 'Invalid object use
End Sub

