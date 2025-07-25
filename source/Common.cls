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

Private Const TRACEABLE_BITMAP_SIZE As Long = 1000
Private Const MM_PER_INCH = 25.4
Private Const MIN_OFFSET As Double = 0.025

'===============================================================================

Public Function TraceBitmap( _
                    ByVal BitmapShape As Shape, _
                    ByVal OrderAbove As Boolean _
                ) As Shape

    Dim ShapeToProcess As Shape
    Dim DeleteShape As Boolean
                
    If BitmapShape.Bitmap.SizeWidth > TRACEABLE_BITMAP_SIZE _
    Or BitmapShape.Bitmap.SizeHeight > TRACEABLE_BITMAP_SIZE Then
        Set ShapeToProcess = BitmapShape.Duplicate
        DeleteShape = True
        Set ShapeToProcess = ShapeToProcess.ConvertToBitmapEx( _
            BitmapShape.Bitmap.Mode, , _
            BitmapShape.Bitmap.Transparent, _
            TRACEABLE_BITMAP_SIZE / (GreaterDim(BitmapShape) / MM_PER_INCH) _
        )
        'тут мы слегка раздуваем битмап, чтобы нормально трассировался
        'если будет впритык к краю - будут куски фона по углам
        With ShapeToProcess.Bitmap.CropEnvelope
            .CopyAssign .Contour(AverageDim(BitmapShape) / 10, cdrContourOutside)
        End With
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
                    ByVal Offset As Double, _
                    ByVal RoundCorners As Boolean _
                 ) As Shape
    
    Dim DeleteShape As Boolean
    If Abs(Offset) < MIN_OFFSET Then Offset = 0
    
    Dim Direction As cdrContourDirection
    If Offset > 0 Then
        Direction = cdrContourOutside
    ElseIf Offset < 0 Then
        Direction = cdrContourInside
    ElseIf Offset = 0 Then
        Set Contour = CreateBoundary(Shape)
        Exit Function
    End If
    
    Dim Cap As cdrContourEndCapType
    Dim Corner As cdrContourCornerType
    If RoundCorners Then
        Cap = cdrContourRoundCap
        Corner = cdrContourCornerRound
    Else
        Cap = cdrContourSquareCap
        Corner = cdrContourCornerMiteredOffsetBevel
    End If
    
    If Shape.Type = cdrGroupShape Then
        Set Shape = CreateBoundary(Shape)
        DeleteShape = True
    End If
    
    Set Contour = _
        Shape.CreateContour( _
            Direction:=Direction, _
            Offset:=Abs(Offset), _
            Steps:=1, _
            EndCapType:=Cap, _
            CornerType:=Corner _
        ).Separate.FirstShape
              
    If DeleteShape Then Shape.Delete
        
End Function

Public Sub Smoothen( _
                ByVal Curve As Curve, _
                Optional ByVal FilletAmount As Double = 1 _
            )
    If FilletAmount = 0 Then FilletAmount = Curve.Length / 100
    With Curve.Nodes.All
        .Smoothen 1
        .AutoReduce 1
        .Fillet FilletAmount
        .AutoReduce 1
        .Fillet FilletAmount / 2
        .AutoReduce 1
    End With
End Sub

Public Sub AverageFill( _
               ByVal SourceShape As Shape, _
               ByVal TargetShape As Shape _
           )
    If SourceShape.Type = cdrGroupShape Then
        TargetShape.Fill.ApplyUniformFill _
            LibCore.GetAverageColorFromShapes( _
                Shapes:=SourceShape.Shapes.All, Fills:=True, Outlines:=False _
            )
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

Private Function Trace(ByVal BitmapShape As Shape) As Shape
    With BitmapShape.Bitmap.Trace(cdrTraceLineArt)
        .BackgroundRemovalMode = cdrTraceBackgroundAutomatic
        .RemoveBackground = True
        Set Trace = .Finish.Group
    End With
End Function

Public Sub ThrowConfigError()
    VBA.Err.Raise CustomError, , LocalizedStrings("Common.ErrConfig")
End Sub

Private Sub Class_Initialize()
    If Not Me Is Common Then Err.Raise 425 'Invalid object use
End Sub
