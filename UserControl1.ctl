VERSION 5.00
Begin VB.UserControl OnFormMenu 
   BackColor       =   &H00C0FFC0&
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
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private m_ParentHwnd As Long
Public Sub ShowMenu()
    If m_ParentHwnd <> 0 Then Exit Sub
        m_ParentHwnd = SetParent(FrmTela(TelaAtual).hwnd, hwnd)
    
    FrmTela(TelaAtual).Move 0, 0
    FrmTela(TelaAtual).Show
End Sub
Public Sub HideMenu()
    If m_ParentHwnd = 0 Then Exit Sub
        FrmMenu.Hide
        SetParent FrmTela(TelaAtual).hwnd, m_ParentHwnd
        m_ParentHwnd = 0
    Unload FrmMenu
End Sub

