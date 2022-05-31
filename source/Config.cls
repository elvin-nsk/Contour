VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Config v2022.05.22
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public Offset As Double
Public OptionMakeOutline As Boolean
Public OutlineColor As String
Public OutlineWidth As Double
Public OptionMakeFill As Boolean
Public OptionFillColor As Boolean
Public OptionMatchColor As Boolean
Public FillColor As String
Public OptionTrace As Boolean

Public OptionSourceAsOne As Boolean
Public OptionSourceAsIs As Boolean
Public OptionSourceWithinGroups As Boolean

Public OptionResultAbove As Boolean
Public OptionResultBelow As Boolean
Public OptionResultAsObjects As Boolean
Public OptionResultAsGroup As Boolean
Public OptionResultAsLayer As Boolean
Public Name As String

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.Create
    
        iniFilePrepare .Self, Action
        
        .Section = "Contour"
        .KeyVar "Offset", Offset, 3#
        .KeyVar "OptionMakeOutline", OptionMakeOutline, True
        .KeyVar "OutlineColor", OutlineColor, "CMYK,USER,0,0,0,100"
        .KeyVar "OutlineWidth", OutlineWidth, 0.1
        .KeyVar "OptionMakeFill", OptionMakeFill, False
        .KeyVar "OptionFillColor", OptionFillColor, False
        .KeyVar "OptionMatchColor", OptionMatchColor, True
        .KeyVar "FillColor", FillColor, "CMYK,USER,0,0,0,0"
        .KeyVar "OptionTrace", OptionTrace, True
        
        .Section = "Source"
        .KeyVar "OptionSourceAsOne", OptionSourceAsOne, False
        .KeyVar "OptionSourceAsIs", OptionSourceAsIs, True
        .KeyVar "OptionSourceWithinGroups", OptionSourceWithinGroups, False
        
        .Section = "Result"
        .KeyVar "OptionResultAbove", OptionResultAbove, True
        .KeyVar "OptionResultBelow", OptionResultBelow, False
        .KeyVar "OptionResultAsObjects", OptionResultAsObjects, True
        .KeyVar "OptionResultAsGroup", OptionResultAsGroup, False
        .KeyVar "OptionResultAsLayer", OptionResultAsLayer, False
        .KeyVar "Name", Name, "контур"
    
    End With
    
End Sub

'===============================================================================

Public Function Load() As Config
    Set Load = New Config
    Load.ForceLoad
End Function

Private Sub Class_Terminate()
    ForceSave
End Sub

'===============================================================================

Public Sub ForceLoad()
    Assign iniRead
End Sub

Public Sub ForceSave()
    Assign iniWrite
End Sub

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
        End Select
    End With

End Sub

Public Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Public Sub SaveStrToFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
