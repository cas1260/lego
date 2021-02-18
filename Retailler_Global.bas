Attribute VB_Name = "Global"
Option Explicit

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000


Public Declare Function GetWindowLong _
    Lib "user32" _
    Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) _
    As Long

Public Declare Function SetWindowLong _
    Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long


Public Declare Function SetWindowPos _
    Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) _
    As Long

Public Sub Retaillable(ctrl As Control, b As Boolean)
    Dim lngStyle As Long, P As Long
    Dim x As Long
    
    'A) Tu récupères le style de ton contrôle
    lngStyle = GetWindowLong(ctrl.hwnd, GWL_STYLE)
    
    'B)  On modifie le Style
    If b Then
        lngStyle = lngStyle Or WS_THICKFRAME
    Else
        lngStyle = lngStyle Xor WS_THICKFRAME
    End If
    
    'C) On met à Jour
    
    x = SetWindowLong(ctrl.hwnd, GWL_STYLE, lngStyle)
    'X = SetWindowPos(ctrl.hwnd, Form1.hwnd, 0, 0, 0, 0, SWP_FLAGS)
    x = SetWindowPos(ctrl.hwnd, FrmTela(TelaAtual).hwnd, 0, 0, 0, 0, SWP_FLAGS)
End Sub
