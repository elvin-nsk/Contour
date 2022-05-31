Attribute VB_Name = "Contour"
'===============================================================================
'   Макрос          : Contour
'   Версия          : 2022.05.31
'   Сайты           : https://vk.com/elvin_macro/Contour
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

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
    With ActiveLayer
        If Not .Visible Or Not .Editable Then
            VBA.MsgBox "Текущий слой закрыт", vbInformation
            GoTo Finally
        End If
    End With
    
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
        RawShapes.AddRange _
            Shapes.Shapes.FindShapes(Query:="Not @Type = 'Group'")
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
            Shape, Contours, Cfg.OptionMatchColor, Cfg
    Next Shape
    For Each Shape In InvalidShapes
        Set BaseShape = Common.TryMakeBaseShape(Shape)
        If Not BaseShape Is Nothing Then
            TempShapes.Add BaseShape
            ContourAndAddToRange _
                BaseShape, Contours, False, Cfg
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
        Set Shape = lib_elvin.WeldShapes(Contours)
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
                    ByVal AssignFill As Boolean, _
                    ByVal Cfg As Config _
                 )
                 
        Dim TempShape As Shape
        Dim NewContour As Shape
        
        'хак с LinkAsChildOf - вытаскиваем сорс для контура из групп
        'чтобы заработало undo
        If Cfg.OptionResultAsObjects Then
            Set NewContour = Common.Contour(Shape, Cfg.Offset)
            If Cfg.OptionResultAbove Then
                NewContour.OrderFrontOf Shape
            Else
                NewContour.OrderBackOf Shape
            End If
        ElseIf Cfg.OptionSourceWithinGroups Then
            If Not Shape.ParentGroup Is Nothing Then
                Set TempShape = Shape.Duplicate
                TempShape.TreeNode.LinkAsChildOf Shape.Layer.TreeNode
                Set NewContour = Common.Contour(TempShape, Cfg.Offset)
                TempShape.Delete
            Else
                Set NewContour = Common.Contour(Shape, Cfg.Offset)
            End If
        Else
            Set NewContour = Common.Contour(Shape, Cfg.Offset)
        End If
        
        If AssignFill Then
            Common.AverageFill Shape, NewContour
        Else
            NewContour.Fill.ApplyNoFill
        End If
        
        ContoursRange.Add NewContour
        
End Sub

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
        ShapeOrShapes.OrderFrontOf lib_elvin.GetTopOrderShape(SourceShapes)
    Else
        ShapeOrShapes.OrderBackOf lib_elvin.GetBottomOrderShape(SourceShapes)
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

Private Sub testShapeParent()
    Debug.Print ActiveShape.ParentGroup Is Nothing
End Sub
