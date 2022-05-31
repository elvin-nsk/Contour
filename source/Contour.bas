Attribute VB_Name = "Contour"
'===============================================================================
'   Макрос          : Contour
'   Версия          : 2022.05.18
'   Сайты           : https://vk.com/elvin_macro/Contour
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = False

Public Const APP_NAME As String = "Contour"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================

Public LocalizedStrings As IStringLocalizer

Sub Start()
    
    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit
    
    If ActiveDocument Is Nothing Then
        VBA.MsgBox "Нет активного документа", vbCritical
        GoTo Finally
    End If
    Dim Source As ShapeRange
    Set Source = ActiveSelectionRange
    If Source.Count = 0 Then
        VBA.MsgBox "Выделите объекты", vbInformation
        GoTo Finally
    End If
    ActiveDocument.Unit = cdrMillimeter
    
    Dim Cfg As Config
    Set Cfg = Config.Load
    If Not ShowViewAndGetResult(Cfg) Then GoTo Finally
    
    lib_elvin.BoostStart APP_NAME, RELEASE
    
    Main Source, Cfg
    
    Source.CreateSelection

Finally:
    lib_elvin.BoostFinish
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
    
    Dim RawShapes As ShapeRange
    Set RawShapes = CreateShapeRange
    
    If Cfg.OptionSourceWithinGroups Then
        RawShapes.AddRange Shapes.Shapes.FindShapes
    Else
        RawShapes.AddRange Shapes
    End If
    
    Dim ReadyShapes As ShapeRange
    Dim InvalidShapes As ShapeRange
    Set InvalidShapes = Common.SeparateInvalidForContour(RawShapes)
    Set ReadyShapes = RawShapes
    
    Dim TempShapes As ShapeRange
    Set TempShapes = CreateShapeRange

    If Cfg.OptionTrace Then
        TempShapes.AddRange _
            Common.ExcludeAndTraceBitmaps(ReadyShapes, Cfg.OptionResultAbove)
        ReadyShapes.AddRange TempShapes
    End If
    
    Dim Contours As ShapeRange
    Set Contours = CreateShapeRange

    Dim Shape As Shape
    
    Dim BaseShape As Shape
    Dim Contour As Shape
    For Each Shape In ReadyShapes
        ContourAndAddToRange _
            Shape, Contours, _
            Cfg.Offset, Cfg.OptionResultAbove, Cfg.OptionMatchColor
    Next Shape
    For Each Shape In InvalidShapes
        Set BaseShape = Common.TryMakeBaseShape(Shape)
        If Not BaseShape Is Nothing Then
            TempShapes.Add BaseShape
            ContourAndAddToRange _
                BaseShape, Contours, _
                Cfg.Offset, Cfg.OptionResultAbove, False
        End If
    Next Shape
    
    If Contours.Count = 0 Then Exit Sub
    
    Dim OutlineColor As Color
    Dim FillColor As Color
    Set OutlineColor = CreateColor(Cfg.OutlineColor)
    Set FillColor = CreateColor(Cfg.FillColor)
    For Each Shape In Contours
        If Cfg.OptionMakeOutline Then
            Shape.Outline.SetProperties Cfg.OutlineWidth
            Shape.Outline.Color.CopyAssign OutlineColor
        Else
            Shape.Outline.SetNoOutline
        End If
        If Cfg.OptionMakeFill Then
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
    
    Dim AverageColor As Color
    If Cfg.OptionSourceAsOne Then
        If Cfg.OptionMatchColor Then
            Set AverageColor = _
                lib_elvin.GetAverageColorFromShapesFill(Contours)
        End If
        Set Shape = WeldShapes(Contours)
        If AverageColor Is Nothing Then
            Shape.Fill.ApplyNoFill
        Else
            Shape.Fill.ApplyUniformFill AverageColor
        End If
        NameAndOrderShape Shape, Shapes, Cfg
        Set Contours = _
            Common.CreateShapeRangeFromShape(Shape)
    End If
        
    If Cfg.OptionResultAsGroup Then
        If Contours.Count = 1 Then
            Set Shape = Contours.FirstShape
        Else
            Set Shape = Contours.Group
        End If
        NameAndOrderShape Shape, Shapes, Cfg
    ElseIf Cfg.OptionResultAsLayer Then
        lib_elvin.MoveToLayer _
            Contours, _
            Common.GetContourLayer(Cfg.Name, Cfg.OptionResultAbove)
    End If
    
    TempShapes.Delete

End Sub

Private Sub ContourAndAddToRange( _
                    ByVal Shape As Shape, _
                    ByVal ContoursRange As ShapeRange, _
                    ByVal Offset As Double, _
                    ByVal OrderAbove As Boolean, _
                    ByVal AssignFill As Boolean _
                 )
        
        Dim NewContour As Shape
        Set NewContour = Common.Contour(Shape, Offset)
        
        If OrderAbove Then
            NewContour.OrderFrontOf Shape
        Else
            NewContour.OrderBackOf Shape
        End If
        
        If AssignFill Then
            Common.AverageFill Shape, NewContour
        Else
            NewContour.Fill.ApplyNoFill
        End If
        
        ContoursRange.Add NewContour
        
End Sub

Private Function WeldShapes(ByVal Shapes As ShapeRange) As Shape
    Set WeldShapes = Shapes.FirstShape
    WeldShapes.CreateSelection
    Dim Shape1 As Shape
    Dim Shape2 As Shape
    Do Until Shapes.Count = 1
        Set Shape1 = Shapes(1)
        Set Shape2 = Shapes(2)
        Shapes.Remove 1
        Shapes.Remove 1
        Set WeldShapes = Shape1.Weld(Shape2)
        Shapes.Add WeldShapes
    Loop
End Function

Private Sub NameAndOrderShape( _
                ByVal Shape As Shape, _
                ByVal SourceShapes As ShapeRange, _
                ByVal Cfg As Config _
            )
    If Cfg.OptionResultAbove Then
        Shape.OrderFrontOf lib_elvin.GetTopOrderShape(SourceShapes)
    Else
        Shape.OrderBackOf lib_elvin.GetBottomOrderShape(SourceShapes)
    End If
    Shape.Name = Cfg.Name
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
    With StringLocalizer.Builder(cdrEnglishUS, New LocalizedStringsRU)
        '.WithLocale cdrRussian, New LocalizedStringsRU
        Set LocalizedStrings = .Build
    End With
End Sub

'===============================================================================
' тесты
'===============================================================================

Private Sub testTraceBitmaps()
    ActiveDocument.Unit = cdrMillimeter
    Common.TraceBitmaps ActiveSelectionRange, True
End Sub

Private Sub testContour()
    ActiveDocument.BeginCommandGroup "testContour"
    ActiveDocument.Unit = cdrMillimeter
    Common.Contour(ActiveShape, 3).Outline.Color.CMYKAssign 0, 0, 0, 100
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
