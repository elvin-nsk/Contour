VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrRussian (1049)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
        
        .Add "Common.ErrNoDocument", "Нет активного документа"
        .Add "Common.ErrNoSelection", "Выделите объекты"
        .Add "Common.ErrDisabledLayer", _
            "Активный слой заблокирован или невидим"
        .Add "Common.ErrConfig", "Ошибка в настройках"
        
        .Add "ProgressBar.TraceCaption", "Трассировка..."
        .Add "ProgressBar.TraceMiddle", "из"
        
        .Add "MainView.Caption", "Тверской, мать его, контур"
        .Add "MainView.FrameContour", "Контур"
        .Add "MainView.LabelOffset", "Отступ"
        .Add "MainView.LabelOffsetUnits", "мм"
        .Add "MainView.OptionMakeOutline", "Абрис"
        .Add "MainView.LabelOutlineWidth", "Толщина абриса"
        .Add "MainView.LabelOutlineUnits", "мм"
        .Add "MainView.OptionMakeFill", "Заливка"
        .Add "MainView.OptionMatchColor", "Подобрать цвет"
        .Add "MainView.OptionTrace", "Трассировать растр"
        .Add "MainView.OptionRoundCorners", "Сглаживать"
        
        .Add "MainView.FrameSource", "Обрабатывать выделенное"
        .Add "MainView.OptionSourceAsOne", "как единый объект"
        .Add "MainView.OptionSourceAsIs", "каждый объект отдельно"
        .Add "MainView.OptionSourceWithinGroups", _
            "каждый объект отдельно, сгруппированные считать отдельными объектами"
        
        .Add "MainView.FrameResult", "Результат"
        .Add "MainView.OptionResultAbove", "над"
        .Add "MainView.OptionResultBelow", "под"
        .Add "MainView.OptionResultAsObjects", "каждым объектом"
        .Add "MainView.OptionResultAsGroup", "объектами отдельной группой"
        .Add "MainView.OptionResultAsLayer", "текущим слоем на отдельном слое"
        .Add "MainView.LabelName", "Имя"
        .Add "MainView.TextBoxName.Default", "контур"
        
        .Add "MainView.ButtonOk", "Пуск"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
