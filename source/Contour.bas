Attribute VB_Name = "Contour"
'===============================================================================
'   Макрос          : Contour
'   Версия          : 2025.06.19
'   Сайты           : https://vk.com/elvin_macro/Contour
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Contour"
Public Const APP_URL As String = "https://vk.com/tverlogo"
Public Const APP_VERSION As String = "2025.06.19"

'===============================================================================
' # Entry points

Public LocalizedStrings As IStringLocalizer

Sub Start()
    
    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit
    
    Dim Source As ShapeRange
    With InputData.GetShapes( _
                       ErrMsgNoDocument:= _
                           LocalizedStrings("Common.ErrNoDocument"), _
                       LayerMustBeEnabled:=True, _
                       ErrMsgLayerDisabled:= _
                           LocalizedStrings("Common.ErrDisabledLayer"), _
                       ErrNoSelection:= _
                           LocalizedStrings("Common.ErrNoSelection") _
                   )
        If .IsError Then
            GoTo Finally
        Else
            Set Source = .Shapes
        End If
    End With
    
    Dim Cfg As Config
    Set Cfg = Config.Load
    If Not ShowViewAndGetResult(Cfg) Then GoTo Finally
    
    LibCore.BoostStart APP_NAME, RELEASE
    
    Main Source, Cfg
    
    Source.CreateSelection

Finally:
    LibCore.BoostFinish
    Set Cfg = Nothing
    Set LocalizedStrings = Nothing
    Exit Sub

Catch:
    MsgBox "Ошибка: " & Err.Description, vbCritical
    Resume Finally

End Sub

Private Sub Main( _
                ByVal Shapes As ShapeRange, _
                ByVal Cfg As Config _
            )
    
    Dim RawShapes As ShapeRange: Set RawShapes = CreateShapeRange
    
    If Cfg.OptionSourceWithinGroups Then
        RawShapes.AddRange _
            Shapes.Shapes.FindShapes(Query:="Not @Type = 'Group'")
    Else
        RawShapes.AddRange Shapes
    End If
    
    Dim ReadyShapes As ShapeRange
    Dim InvalidShapes As ShapeRange: Set InvalidShapes = _
        Common.SeparateInvalidForContour(RawShapes)
    Set ReadyShapes = RawShapes
    
    Dim TempShapes As ShapeRange: Set TempShapes = CreateShapeRange

    If Cfg.OptionTrace Then
        TempShapes.AddRange _
            ExcludeAndTraceBitmaps( _
                ReadyShapes, Cfg.OptionResultAbove _
            )
        ReadyShapes.AddRange TempShapes
    End If
    
    Dim Contours As ShapeRange: Set Contours = CreateShapeRange
    Dim Contours2 As ShapeRange: Set Contours2 = CreateShapeRange

    Dim Shape As Shape

    Dim BaseShape As Shape
    Dim Contour As Shape
    For Each Shape In ReadyShapes
        ContourAndAddToRange _
            Shape, Contours, Contours2, Cfg.OptionMatchColor, Cfg
    Next Shape
    For Each Shape In InvalidShapes
        Set BaseShape = Common.TryMakeBaseShape(Shape)
        If Not BaseShape Is Nothing Then
            TempShapes.Add BaseShape
            ContourAndAddToRange _
                BaseShape, Contours, Contours2, False, Cfg
        End If
    Next Shape
    
    If Contours.Count = 0 Then Exit Sub
    
    ProcessContours Contours, True, Cfg
    If Contours2.Count > 0 Then ProcessContours Contours2, False, Cfg
    
    Dim AverageColor As Color
    If Cfg.OptionSourceAsOne Then
        If Cfg.OptionMatchColor Then
            Set AverageColor = _
                GetAverageColorFromShapes( _
                    Shapes:=Contours, Fills:=True, Outlines:=False _
                )
        End If
        Set Shape = Weld(Contours)
        If Not AverageColor Is Nothing Then
            Shape.Fill.ApplyUniformFill AverageColor
        End If
        NameAndOrderShape Shape, Shapes, Cfg
        Set Contours = _
            Common.CreateShapeRangeFromShape(Shape)
        If Contours2.Count > 0 Then
            Set Shape = Weld(Contours2)
            Shape.Name = Cfg.Name
            Shape.OrderBackOf Contours.FirstShape
        End If
    End If
        
    Dim ContourLayer As Layer
    If Cfg.OptionResultAsGroup Then
        If Contours.Count = 1 Then
            Set Shape = Contours.FirstShape
        Else
            Set Shape = Contours.Group
        End If
        NameAndOrderShape Shape, Shapes, Cfg
    ElseIf Cfg.OptionResultAsLayer Then
        Set ContourLayer = _
            Common.GetContourLayer(Cfg.Name, Cfg.OptionResultAbove)
        Contours2.MoveToLayer ContourLayer
        Contours.MoveToLayer ContourLayer
    End If
    
    TempShapes.Delete

End Sub

Private Sub ProcessContours( _
                ByVal Contours As ShapeRange, _
                ByVal AssignFill As Boolean, _
                ByVal Cfg As Config _
            )
    Dim OutlineColor As Color: Set OutlineColor = CreateColor(Cfg.OutlineColor)
    Dim FillColor As Color: Set FillColor = CreateColor(Cfg.FillColor)
    Dim Shape As Shape
    For Each Shape In Contours
        If Cfg.OptionMakeOutline Then
            Shape.Outline.SetProperties Cfg.OutlineWidth
            Shape.Outline.Color.CopyAssign OutlineColor
        Else
            Shape.Outline.SetNoOutline
        End If
        If Cfg.OptionMakeFill And AssignFill Then
            If Cfg.OptionFillColor Then
                Shape.Fill.ApplyUniformFill FillColor
            End If
        Else
            Shape.Fill.ApplyNoFill
        End If
        If Not Cfg.OptionSourceAsOne Then
            Shape.Name = Cfg.Name
        End If
    Next Shape
End Sub

Private Function ExcludeAndTraceBitmaps( _
                    ByVal Shapes As ShapeRange, _
                    ByVal OrderAbove As Boolean _
                ) As ShapeRange
    Set ExcludeAndTraceBitmaps = CreateShapeRange
    Dim BitmapShapes As ShapeRange
    Set BitmapShapes = Shapes.Shapes.FindShapes(Type:=cdrBitmapShape)
    If BitmapShapes.Count > 0 Then
        Dim PBar As IProgressBar
        Set PBar = ProgressBar.CreateNumeric(BitmapShapes.Count)
        PBar.Caption = LocalizedStrings("ProgressBar.TraceCaption")
        PBar.NumericMiddleText = LocalizedStrings("ProgressBar.TraceMiddle")
        Shapes.RemoveRange BitmapShapes
        ExcludeAndTraceBitmaps.AddRange _
            TraceBitmaps(BitmapShapes, OrderAbove, PBar)
    End If
End Function

Private Function TraceBitmaps( _
                    ByVal BitmapShapes As ShapeRange, _
                    ByVal OrderAbove As Boolean, _
                    ByVal PBar As IProgressBar _
                ) As ShapeRange
    Set TraceBitmaps = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In BitmapShapes
        TraceBitmaps.Add Common.TraceBitmap(Shape, OrderAbove)
        PBar.Update
    Next Shape
End Function

Private Sub ContourAndAddToRange( _
                    ByVal Shape As Shape, _
                    ByVal ContoursRange As ShapeRange, _
                    ByVal ContoursRange2 As ShapeRange, _
                    ByVal AssignFill As Boolean, _
                    ByVal Cfg As Config _
                 )
        Dim TempShape As Shape, DeleteTempShape As Boolean, Reorder As Boolean
        
        If Cfg.OptionResultAsObjects Then
            Set TempShape = Shape
            Reorder = True
        ElseIf Cfg.OptionSourceWithinGroups Then
            If Not Shape.ParentGroup Is Nothing Then
                Set TempShape = Shape.Duplicate
                'хак с LinkAsChildOf - вытаскиваем сорс для контура из групп
                'чтобы работало undo
                TempShape.TreeNode.LinkAsChildOf Shape.Layer.TreeNode
                DeleteTempShape = True
            Else
                Set TempShape = Shape
            End If
        Else
            Set TempShape = Shape
        End If
        
        Dim NewContour As Shape: Set NewContour = _
            MakeContour(TempShape, Cfg.Offset, Cfg.OptionRoundCorners, AssignFill)
        ContoursRange.Add NewContour
        
        If Reorder Then
            If Cfg.OptionResultAbove Then
                NewContour.OrderFrontOf Shape
            Else
                NewContour.OrderBackOf Shape
            End If
        End If
        
        Dim NewContour2 As Shape
        If Cfg.OptionSecondaryContour Then
            Set NewContour2 = _
                MakeContour( _
                    TempShape, Cfg.Offset2, Cfg.OptionRoundCorners2, False _
                )
            ContoursRange2.Add NewContour2
            NewContour2.OrderBackOf NewContour
        End If
        
        If DeleteTempShape Then TempShape.Delete
        
End Sub

Private Function MakeContour( _
                    ByVal Shape As Shape, _
                    ByVal Offset As Double, _
                    ByVal RoundCorners As Boolean, _
                    ByVal AssignFill As Boolean _
                 ) As Shape
    Dim FilletAmount As Double: FilletAmount = Abs(Offset)
    
    Set MakeContour = _
            Common.Contour(Shape, Offset, RoundCorners)
    If RoundCorners Then
        MakeContour.ConvertToCurves
        Common.Smoothen MakeContour.Curve, FilletAmount
    End If
    If AssignFill Then
        Common.AverageFill Shape, MakeContour
    Else
        MakeContour.Fill.ApplyNoFill
    End If
End Function

Private Sub NameAndOrderShape( _
                ByVal Shape As Shape, _
                ByVal SourceShapes As ShapeRange, _
                ByVal Cfg As Config _
            )
    OrderShapeOrShapes Shape, SourceShapes, Cfg
    Shape.Name = Cfg.Name
End Sub

Private Sub OrderShapeOrShapes( _
                ByVal ShapeOrShapes As Object, _
                ByVal SourceShapes As ShapeRange, _
                ByVal Cfg As Config _
            )
    If Cfg.OptionResultAbove Then
        ShapeOrShapes.OrderFrontOf LibCore.GetTopOrderShape(SourceShapes)
    Else
        ShapeOrShapes.OrderBackOf LibCore.GetBottomOrderShape(SourceShapes)
    End If
End Sub

Private Function ShowViewAndGetResult(ByVal Cfg As Config) As Boolean
    With New MainView
    
        .OffsetHandler = Cfg.Offset
        .OptionMakeOutline = Cfg.OptionMakeOutline
        Set .OutlineColor = CreateColor(Cfg.OutlineColor)
        .OutlineWidthHandler = Cfg.OutlineWidth
        .OptionMakeFill = Cfg.OptionMakeFill
        .OptionFillColor = Cfg.OptionFillColor
        .OptionMatchColor = Cfg.OptionMatchColor
        Set .FillColor = CreateColor(Cfg.FillColor)
        .OptionTrace = Cfg.OptionTrace
        .OptionRoundCorners = Cfg.OptionRoundCorners
        
        .OptionSecondaryContour = Cfg.OptionSecondaryContour
        .TextBoxOffset2 = Cfg.Offset2
        .OptionRoundCorners2 = Cfg.OptionRoundCorners2
        
        .OptionSourceAsOne = Cfg.OptionSourceAsOne
        .OptionSourceAsIs = Cfg.OptionSourceAsIs
        .OptionSourceWithinGroups = Cfg.OptionSourceWithinGroups
        
        .OptionResultAbove = Cfg.OptionResultAbove
        .OptionResultBelow = Cfg.OptionResultBelow
        .OptionResultAsObjects = Cfg.OptionResultAsObjects
        .OptionResultAsGroup = Cfg.OptionResultAsGroup
        .OptionResultAsLayer = Cfg.OptionResultAsLayer
        .NameHandler = Cfg.Name
        
        .Show
        ShowViewAndGetResult = .IsOk
        If Not .IsOk Then Exit Function
        
        Cfg.Offset = .OffsetHandler
        Cfg.OptionMakeOutline = .OptionMakeOutline
        Cfg.OutlineColor = .OutlineColor.ToString
        Cfg.OutlineWidth = .OutlineWidthHandler
        Cfg.OptionMakeFill = .OptionMakeFill
        Cfg.OptionFillColor = .OptionFillColor
        Cfg.OptionMatchColor = .OptionMatchColor
        Cfg.FillColor = .FillColor.ToString
        Cfg.OptionTrace = .OptionTrace
        Cfg.OptionRoundCorners = .OptionRoundCorners
        
        Cfg.OptionSecondaryContour = .OptionSecondaryContour
        Cfg.Offset2 = .OffsetHandler2
        Cfg.OptionRoundCorners2 = .OptionRoundCorners2
        
        Cfg.OptionSourceAsOne = .OptionSourceAsOne
        Cfg.OptionSourceAsIs = .OptionSourceAsIs
        Cfg.OptionSourceWithinGroups = .OptionSourceWithinGroups
        
        Cfg.OptionResultAbove = .OptionResultAbove
        Cfg.OptionResultBelow = .OptionResultBelow
        Cfg.OptionResultAsObjects = .OptionResultAsObjects
        Cfg.OptionResultAsGroup = .OptionResultAsGroup
        Cfg.OptionResultAsLayer = .OptionResultAsLayer
        Cfg.Name = .NameHandler
    
    End With
End Function

'===============================================================================

Private Sub LocalizedStringsInit()
    With StringLocalizer.Builder(cdrEnglishUS, New LocalizedStringsEN)
        .WithLocale cdrRussian, New LocalizedStringsRU
        Set LocalizedStrings = .Build
    End With
End Sub

'===============================================================================
' тесты
'===============================================================================

Private Sub testTraceBitmaps()
    ActiveDocument.Unit = cdrMillimeter
    Dim PBar As IProgressBar
    Set PBar = ProgressBar.CreateNumeric(ActiveSelectionRange.Count)
    TraceBitmaps ActiveSelectionRange, True, PBar
End Sub

Private Sub testContour()
    ActiveDocument.BeginCommandGroup "testContour"
    ActiveDocument.Unit = cdrMillimeter
    Common.Contour(ActiveShape, 3, True).Outline.Color.CMYKAssign 0, 0, 0, 100
    ActiveDocument.EndCommandGroup
End Sub

Private Sub testWeld()
    ActiveSelectionRange.FirstShape.Weld ActiveSelectionRange.LastShape
End Sub

Private Sub testBlend()
    ActiveSelectionRange.LastShape.Fill.UniformColor.BlendWith _
        ActiveSelectionRange.LastShape.Fill.UniformColor, 50
End Sub

Private Sub testZOrder()
    GetBottomOrderShape(ActiveSelectionRange).CreateSelection
End Sub

