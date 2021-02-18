Attribute VB_Name = "NoClose"
Option Explicit

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Private ReadyToClose As Boolean
Private Function RemoveMenus(Frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(Frm.hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Function

Public Sub Desab(Fecha1 As Boolean, Max1 As Boolean, Min1 As Boolean, Frm As Form)
    RemoveMenus Frm, False, False, _
        False, Not Min1, Not Max1, True, Not Fecha1
End Sub





