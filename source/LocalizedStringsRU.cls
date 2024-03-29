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
        
        .Add "Common.ErrNoDocument", "��� ��������� ���������"
        .Add "Common.ErrNoSelection", "�������� �������"
        .Add "Common.ErrDisabledLayer", _
            "�������� ���� ������������ ��� �������"
        .Add "Common.ErrConfig", "������ � ����������"
        
        .Add "ProgressBar.TraceCaption", "�����������..."
        .Add "ProgressBar.TraceMiddle", "��"
        
        .Add "MainView.Caption", "��������, ���� ���, ������"
        .Add "MainView.FrameContour", "������"
        .Add "MainView.LabelOffset", "������"
        .Add "MainView.LabelOffsetUnits", "��"
        .Add "MainView.OptionMakeOutline", "�����"
        .Add "MainView.LabelOutlineWidth", "������� ������"
        .Add "MainView.LabelOutlineUnits", "��"
        .Add "MainView.OptionMakeFill", "�������"
        .Add "MainView.OptionMatchColor", "��������� ����"
        .Add "MainView.OptionTrace", "������������ �����"
        .Add "MainView.OptionRoundCorners", "��������� ����"
        
        .Add "MainView.FrameSource", "������������ ����������"
        .Add "MainView.OptionSourceAsOne", "��� ������ ������"
        .Add "MainView.OptionSourceAsIs", "������ ������ ��������"
        .Add "MainView.OptionSourceWithinGroups", _
            "������ ������ ��������, ��������������� ������� ���������� ���������"
        
        .Add "MainView.FrameResult", "���������"
        .Add "MainView.OptionResultAbove", "���"
        .Add "MainView.OptionResultBelow", "���"
        .Add "MainView.OptionResultAsObjects", "������ ��������"
        .Add "MainView.OptionResultAsGroup", "��������� ��������� �������"
        .Add "MainView.OptionResultAsLayer", "������� ����� �� ��������� ����"
        .Add "MainView.LabelName", "���"
        .Add "MainView.TextBoxName.Default", "������"
        
        .Add "MainView.ButtonOk", "����"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
