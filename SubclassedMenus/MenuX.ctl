VERSION 5.00
Begin VB.UserControl MenuX 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1320
   ScaleWidth      =   1545
   Begin VB.Image I 
      Height          =   480
      Left            =   240
      Picture         =   "MenuX.ctx":0000
      Stretch         =   -1  'True
      Top             =   420
      Width           =   480
   End
End
Attribute VB_Name = "MenuX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim hMenu As Long
Dim WithEvents myMnu As Menu
Attribute myMnu.VB_VarHelpID = -1
Dim hMenuWork As Long
Dim hMenuSub As Long
Dim hMenuItems As Long
Dim result As Long
Dim MenuXX()  As Long

Dim ContMenu As Long
Dim MenuPrincipal As Long
Dim SubMenuP As Long
Public FormX As Form
Public Event Click(Nome As String)

Private Sub UserControl_Initialize()
ContMenu = 0
MenuPrincipal = 0
SubMenuP = 0
End Sub

Private Sub UserControl_Resize()
I.Top = 0
I.Left = 0
I.Height = UserControl.Height
I.Width = UserControl.Width
End Sub

Public Function myAddMenuItem(mnuType As Long, mnuId As Long, hMnu As Long, mnuPos As Long, mnuPosType As Long, mnuStr As String) As Long
'MenuItemInfo with SetMenuItemInfo() also gives the ability to add bitmaps to menu items and check items.

    Dim mii As MENUITEMINFO

    With mii
        ' The size of this structure.
        .cbSize = Len(mii)
        ' Which elements of the structure to use.
        .fMask = MIIM_ID Or MIIM_DATA Or MIIM_TYPE Or MIIM_SUBMENU
        ' The type of item: a string.
        .fType = mnuType
        ' This item is currently enabled and is the default item.
        .fState = MFS_ENABLED Or MFS_DEFAULT
        ' Assign this item an item identifier.
        .wID = mnuId
        ' Display the following text for the item.
        .dwTypeData = mnuStr
        .cch = Len(.dwTypeData)
    End With

    myAddMenuItem = InsertMenuItem(hMnu, mnuPos, mnuPosType, mii)
    
End Function

Public Function AddMenu(nMenu As String, Index As Long, Optional SubMenu As Boolean)
    
    If Index = 0 Then
        hMenu = GetMenu(FormX.hwnd)
        Menu0 = CreatePopupMenu
        MenuXX(Index) = Menu0
        If SubMenu = True Then
            result = AppendMenu(hMenu, 0, Menu0, nMenu)
        Else
            result = AppendMenu(hMenu, MF_POPUP, Menu0, nMenu)
        End If
        'MenuPrincipal = MenuPrincipal + 1
    Else
        MenuXX(Index) = CreatePopupMenu
        If SubMenu = True Then
            result = AppendMenu(MenuXX(Index - 1), MF_STRING, MenuXX(Index), nMenu)
            'result = myAddMenuItem(MF_STRING, mnuFileAnmlDog, hMenuWork, 0, 1, "Dogs")
            'result = myAddMenuItem(MF_STRING, ContMenu, MenuXX(Index), ContMenu, 1, nMenu)
            ContMenu = ContMenu + 1
        Else
            result = AppendMenu(MenuXX(Index - 1), MF_POPUP, MenuXX(Index), nMenu)
        End If
    End If
    
End Function

Public Function FimMenu()

    ProcOld = SetWindowLong(FormX.hwnd, GWL_WNDPROC, AddressOf ClickMenu)

End Function

Public Sub InicializaMenu(Frm As Form)
ReDim MenuXX(50) As Long
Set FormX = Frm
End Sub
