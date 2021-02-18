VERSION 5.00
Begin VB.UserControl OnFormMenu 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   1680
   ScaleWidth      =   4800
End
Attribute VB_Name = "OnFormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private m_ParentHwnd As Long
Public Sub Chama(Tela As Form)
    If m_ParentHwnd <> 0 Then Exit Sub
        m_ParentHwnd = SetParent(Tela.hwnd, hwnd)
    Tela.Move -50, -40
    Tela.Show
End Sub
Public Sub TiraTela(Tela As Form)
    If m_ParentHwnd = 0 Then Exit Sub
        Tela.Hide
        SetParent Tela.hwnd, m_ParentHwnd
        m_ParentHwnd = 0
    Unload Tela
End Sub


Private Sub UserControl_Resize()
'If TelaAtual <> -1 Then
'    With Menus(FrmTela(TelaAtual).Cont)
'        .Top = 0
'        .Left = 0
'        .Width = UserControl.ScaleWidth
'        .Height = UserControl.ScaleHeight
'    End With
'End If
End Sub
