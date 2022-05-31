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
        
        .Add "MainView.frForm", "Формирование композиций"
        .Add "MainView.obByCount", "по количеству"
        .Add "MainView.lbColumns", "столбцов"
        .Add "MainView.lbRows", "строк"
        .Add "MainView.opByArea", "по площади"
        .Add "MainView.lbWidth", "мм по ширине"
        .Add "MainView.lbHeight", "мм по высоте"
        .Add "MainView.opByPage", "по размеру страницы"
        
        .Add "MainView.frCompose", "Расположение композиций"
        .Add "MainView.opOnePage", "на одной странице"
        .Add "MainView.opMultiplePages", "каждая на своей странице"
        
        .Add "MainView.frSpaces", "Расстояния между элементами"
        .Add "MainView.lbHorizontalSpace", "мм по ширине"
        .Add "MainView.lbVerticalSpace", "мм по высоте"
        
        .Add "MainView.cbNewDocument", "Создать новый документ"
        .Add "MainView.cbGroup", "Группировать композиции"
        
        .Add "MainView.btnOk", "Пуск"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
