Attribute VB_Name = "Desabilitar"

Option Explicit
Private Declare Function apiGetSystemMenu Lib "user32" Alias "GetSystemMenu" (ByVal HWnd As Long, ByVal bRevert As Long) As Long 'used for dealing with menus
Private Declare Function apiGetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hSysMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MenuItemInfo) As Long
Private Declare Function apiSetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hSysMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpcMenuItemInfo As MenuItemInfo) As Long

Private Type MenuItemInfo
   cbSize As Long
   fMask As eMENUITEMINFOMask
   fType As eMenuType
   fState As eMenuState
   wID As eSysMenuID
   'wID As Long is changed for system  menu
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
   'hbmpItem As eMENUITEMINFOBitmap
   'Windows 98, Windows 2000: Handle to the bitmap to be displayed.
   'It is used when the MIIM_BITMAP flag is set in the fMask member.
End Type

Private Enum eMENUITEMINFOMask   'MENUITEMINFO, UINT fMask (Members to retrieve or set. )
   MIIM_BITMAP = &H80&
   'Windows 98, Windows 2000: Retrieves or sets the hbmpItem member.
   MIIM_CHECKMARKS = &H8&
   'Retrieves or sets the hbmpChecked and hbmpUnchecked members.
   MIIM_DATA = &H20&
   'Retrieves or sets the dwItemData member.
   MIIM_FTYPE = &H100&
   'Windows 98, Windows 2000: Retrieves or sets the fType member.
   MIIM_ID = &H2&
   'Retrieves or sets the wID member.
   MIIM_STATE = &H1&
   'Retrieves or sets the fState member.
   MIIM_STRING = &H40&
   'Windows 98, Windows 2000: Retrieves or sets the dwTypeData member.
   MIIM_SUBMENU = &H4&
   'Retrieves or sets the hSubMenu member.
   MIIM_TYPE = &H10&
   'Retrieves or sets the fType and dwTypeData members. Windows
   '98, Windows 2000: MIIM_TYPE is replaced by MIIM_BITMAP,
   'MIIM_FTYPE and MIIM_STRING.
End Enum 'eMENUITEMINFOMask

Private Enum eMenuType   'MENUITEMINFO, UINT fType (Menu item type.)
   MFT_BITMAP = &H4&
   'Displays the menu item using a bitmap. The low-order word of
   'the dwTypeData member is the bitmap handle, and the
   'cch member is ignored. Windows 98, Windows 2000: MFT_BITMAP
   'is replaced by MIIM_BITMAP and hbmpItem
   MFT_MENUBARBREAK = &H20&
   'Places the menu item on a new line (for a menu bar) or in a
   'new column (for a drop-down menu, submenu, or shortcut
   'menu). For a drop-down menu, submenu, or shortcut
   'menu, a vertical line separates the new column from
   'the old.
   MFT_MENUBREAK = &H40&
   'Places the menu item on a new line (for a menu bar) or in a
   'new column (for a drop-down menu, submenu, or shortcut
   'menu). For a drop-down menu, submenu, or shortcut
   'menu, the columns are not separated by a vertical
   'line.
   MFT_OWNERDRAW = &H100&
   'Assigns responsibility for drawing the menu item to the window
   'that owns the menu. The window receives a WM_MEASUREITEM
   'message before the menu is displayed for the first
   'time, and a WM_DRAWITEM message whenever the appearance
   'of the menu item must be updated. If this value is
   'specified, the dwTypeData member contains an application-defined
   'value.
   MFT_RADIOCHECK = &H200&
   'Displays selected menu items using a radio-button mark instead
   'of a check mark if the hbmpChecked member is NULL.
   MFT_RIGHTJUSTIFY = &H4000&
   'Right-justifies the menu item and any subsequent items. This
   'value is valid only if the menu item is in a menu
   'bar.
   MFT_RIGHTORDER = &H2000&
   'Windows 95/98, Windows 2000: Specifies that menus cascade right-to-left
   '(the default is left-to-right). This is used to support
   'right-to-left languages, such as Arabic and Hebrew.
   '`
   MFT_SEPARATOR = &H800&
   'Specifies that the menu item is a separator. A menu item separator
   'appears as a horizontal dividing line. The dwTypeData
   'and cch members are ignored. This value is valid only
   'in a drop-down menu, submenu, or shortcut menu.
   MFT_STRING = &H0&
   'Displays the menu item using a text string. The dwTypeData member
   'is the pointer to a null-terminated string, and the
   'cch member is the length of the string. Windows 98,
   'Windows 2000: MFT_STRING is replaced by MIIM_STRING
End Enum 'eMenuType

Public Enum eMenuState   'MENUITEMINFO, UINT fState (Menu item state.)
   MFS_CHECKED = &H8&
   'Checks the menu item. For more information about selected menu
   'items, see the hbmpChecked member.
   MFS_DEFAULT = &H1000&
   'Specifies that the menu item is the default. A menu can contain
   'only one default menu item, which is displayed in
   'bold.
   MFS_DISABLED = &H3&
   'Disables the menu item and grays it so that it cannot be selected.
   'This is equivalent to MFS_GRAYED.
   MFS_ENABLED = &H0&
   'Enables the menu item so that it can be selected. This is the
   'default state.
   MFS_GRAYED = &H3&
   'Disables the menu item and grays it so that it cannot be selected.
   'This is equivalent to MFS_DISABLED.
   MFS_HILITE = &H80&
   'Highlights the menu item.
   MFS_UNCHECKED = &H0&
   'Unchecks the menu item. For more information about clear menu
   'items, see the hbmpUnchecked member.
   MFS_UNHILITE = &H0&
   'Removes the highlight from the menu item. This is the default
   'state.
End Enum 'eMenuState

Public Enum eSysMenuID 'System Menu ID
   'this values are related to the WM_SYSCOMMAND message
   'and eWM_SYSCOMMAND eunumerator
   SMSC_CLOSE = &HF060&
   SMSC_DEFAULT = &HF160&
   SMSC_MAXIMIZE = &HF030&
   SMSC_MINIMIZE = &HF020&
   SMSC_MOVE = &HF010&
   SMSC_SIZE = &HF000&
   SMSC_RESTORE = &HF120&
End Enum 'eSysMenuID

Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As eWindowOffsets, ByVal dwNewLong As eWindowStyle) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWnd As Long, ByVal nIndex As eWindowOffsets) As Long
Private Enum eWindowOffsets
   GWL_EXSTYLE = (-20) 'Sets a new extended window style. For more information, see CreateWindowEx.
   GWL_STYLE = (-16) 'Sets a new window style.
   GWL_WNDPROC = (-4) 'Sets a new address for the window procedure.
   GWL_HINSTANCE = (-6) 'Sets a new application instance handle.
   GWL_ID = (-12) 'Sets a new identifier of the window.
   GWL_USERDATA = (-21) 'Sets the user data associated with the window. This data is intended for use by the application that created the window.
   'Its value is initially zero.
   'The following values are also available when the hWnd parameter identifies a dialog box.
   DWL_DLGPROC = 4 'Sets the new pointer to the dialog box procedure.
   DWL_MSGRESULT = 0 'Sets the return value of a message processed in the dialog box procedure.
   DWL_USER = 8 'Sets new extra information that is private to the application, such as handles or pointers.
End Enum

Private Enum eWindowStyle
   WS_BORDER = &H800000  'CreateWindow, DWORD dwStyle (window style)  Const WS_BORDER As Long = &H800000&
   WS_CAPTION = &HC00000  'CreateWindow, DWORD dwStyle (window style)  Const WS_CAPTION As Long = &HC00000&
   WS_CHILD = &H40000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_CHILD As Long = &H40000000&
   WS_CHILDWINDOW = WS_CHILD 'CreateWindow, DWORD dwStyle (window style)  Const WS_CHILDWINDOW As Long = WS_CHILD
   WS_CLIPCHILDREN = &H2000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_CLIPCHILDREN As Long = &H2000000&
   WS_CLIPSIBLINGS = &H4000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_CLIPSIBLINGS As Long = &H4000000&
   WS_DISABLED = &H8000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_DISABLED As Long = &H8000000&
   WS_DLGFRAME = &H400000  'CreateWindow, DWORD dwStyle (window style)  Const WS_DLGFRAME As Long = &H400000&
   WS_GROUP = &H20000  'CreateWindow, DWORD dwStyle (window style)  Const WS_GROUP As Long = &H20000&
   WS_HSCROLL = &H100000  'CreateWindow, DWORD dwStyle (window style)  Const WS_HSCROLL As Long = &H100000&
   WS_MINIMIZE = &H20000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_MINIMIZE As Long = &H20000000&
   WS_ICONIC = WS_MINIMIZE 'CreateWindow, DWORD dwStyle (window style)  Const WS_ICONIC As Long = WS_MINIMIZE
   WS_MAXIMIZE = &H1000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_MAXIMIZE As Long = &H1000000&
   WS_MAXIMIZEBOX = &H10000  'CreateWindow, DWORD dwStyle (window style)  Const WS_MAXIMIZEBOX As Long = &H10000&
   'WS_MINIMIZE = &H20000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_MINIMIZE As Long = &H20000000&
   WS_MINIMIZEBOX = &H20000  'CreateWindow, DWORD dwStyle (window style)  Const WS_MINIMIZEBOX As Long = &H20000&
   WS_OVERLAPPED = &H0& 'CreateWindow, DWORD dwStyle (window style)  Const WS_OVERLAPPED As Long = &H0&
   WS_POPUP = &H80000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_POPUP As Long = &H80000000&
   WS_SYSMENU = &H80000  'CreateWindow, DWORD dwStyle (window style)  Const WS_SYSMENU As Long = &H80000&
   WS_TABSTOP = &H10000  'CreateWindow, DWORD dwStyle (window style)  Const WS_TABSTOP As Long = &H10000&
   WS_THICKFRAME = &H40000  'CreateWindow, DWORD dwStyle (window style)  Const WS_THICKFRAME As Long = &H40000&
   WS_TILED = WS_OVERLAPPED 'CreateWindow, DWORD dwStyle (window style)  Const WS_TILED As Long = WS_OVERLAPPED
   WS_VISIBLE = &H10000000  'CreateWindow, DWORD dwStyle (window style)  Const WS_VISIBLE As Long = &H10000000&
   WS_VSCROLL = &H200000  'CreateWindow, DWORD dwStyle (window style)  Const WS_VSCROLL As Long = &H200000&
   WS_SIZEBOX = WS_THICKFRAME 'CreateWindow, DWORD dwStyle (window style)  Const WS_SIZEBOX As Long = WS_THICKFRAME
   WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX 'CreateWindow, DWORD dwStyle (window style)  Const WS_OVERLAPPEDWINDOW As Long = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
   WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU 'CreateWindow, DWORD dwStyle (window style)  Const WS_POPUPWINDOW As Long = WS_POPUP Or WS_BORDER Or WS_SYSMENU
   WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW 'CreateWindow, DWORD dwStyle (window style)  Const WS_TILEDWINDOW As Long = WS_OVERLAPPEDWINDOW
End Enum 'eWindowStyle

Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long '(winui)

Public Function ToggleSysMenuEnableDisable( _
   HWnd As Long, _
   ByVal wMenuID As eSysMenuID) As Long

   Dim retVal As Long
   Dim MenuID As Long
   Dim hSysMenu As Long
   Dim udtMII  As MenuItemInfo

   'Const WM_NCACTIVATE  As Long = &H86
   'Const GWL_STYLE As Long = (-16&)
   'Const WS_MAXIMIZEBOX As Long = &H10000
   'Const WS_MINIMIZEBOX As Long = &H20000

   'first try to get the original menu
   hSysMenu = apiGetSystemMenu(HWnd, 0)

   With udtMII
      .cbSize = Len(udtMII)
      .dwTypeData = VBA.String$(255, 0)
      .cch = Len(.dwTypeData)
      .fMask = eMENUITEMINFOMask.MIIM_STATE
      .wID = wMenuID
   End With 'UDTMII
   retVal = apiGetMenuItemInfo(hSysMenu, udtMII.wID, False, udtMII)

   If retVal = 0 Then
      'if failed, it means that we have changed the menu ID

      udtMII.wID = -wMenuID
      retVal = apiGetMenuItemInfo(hSysMenu, udtMII.wID, False, udtMII)

      'To change the system menu item state
      'We need to change the menu ID.
      With udtMII
         MenuID = udtMII.wID
         .wID = wMenuID
         .fMask = eMENUITEMINFOMask.MIIM_ID
      End With 'UDTMII
      retVal = apiSetMenuItemInfo(hSysMenu, MenuID, False, udtMII)

   Else 'NOT RETVAL...

      'To change the system menu item state
      'We need to change the menu ID.
      With udtMII
         MenuID = udtMII.wID
         If .fState <> (.fState Or MFS_GRAYED) Then
            .wID = -(.wID)
         End If
         .fMask = eMENUITEMINFOMask.MIIM_ID
      End With 'UDTMII
      retVal = apiSetMenuItemInfo(hSysMenu, MenuID, False, udtMII)

   End If

   'if failed, restore the original menu id
   If retVal = 0 Then
      udtMII.wID = MenuID

   Else 'NOT RETVAL...

      With udtMII
         'if disabled
         If .fState = (.fState Or eMenuState.MFS_GRAYED) Then
            'Enable it
            .fState = .fState - eMenuState.MFS_GRAYED
         Else 'NOT udtMII.FSTATE...'NOT .FSTATE...
            'if enabled, Disabled it
            .fState = (.fState Or eMenuState.MFS_GRAYED)
         End If

         .fMask = eMENUITEMINFOMask.MIIM_STATE
         retVal = apiSetMenuItemInfo(hSysMenu, udtMII.wID, False, udtMII)

         If retVal = 0 Then
         
            'if failed, retry with custom ID (just negative value of the original menu item ID)
            MenuID = udtMII.wID

            If udtMII.fState = (udtMII.fState Or MFS_GRAYED) Then
               udtMII.wID = -udtMII.wID
            End If

            udtMII.fMask = MIIM_ID
            retVal = apiSetMenuItemInfo(hSysMenu, MenuID, False, udtMII)
            If retVal = 0 Then
               'if failed, restore the original
               udtMII.wID = MenuID
            End If
            
         End If

         If wMenuID = SMSC_MINIMIZE Then
            'if disabled
            If udtMII.fState = (udtMII.fState Or eMenuState.MFS_GRAYED) Then
               'gray the min box
               Call apiSetWindowLong(HWnd, eWindowOffsets.GWL_STYLE, _
                                    apiGetWindowLong(HWnd, eWindowOffsets.GWL_STYLE) Xor eWindowStyle.WS_MINIMIZEBOX)
            Else
               'draw nnormally
               Call apiSetWindowLong(HWnd, eWindowOffsets.GWL_STYLE, _
                                    apiGetWindowLong(HWnd, eWindowOffsets.GWL_STYLE) Or eWindowStyle.WS_MINIMIZEBOX)
            End If
         ElseIf wMenuID = SMSC_MAXIMIZE Then
            'if disabled
            If udtMII.fState = (udtMII.fState Or eMenuState.MFS_GRAYED) Then
               'gray the max box
               Call apiSetWindowLong(HWnd, eWindowOffsets.GWL_STYLE, _
                                    apiGetWindowLong(HWnd, GWL_STYLE) Xor eWindowStyle.WS_MAXIMIZEBOX)
            Else
               'draw normally
               Call apiSetWindowLong(HWnd, eWindowOffsets.GWL_STYLE, _
                                    apiGetWindowLong(HWnd, eWindowOffsets.GWL_STYLE) Or eWindowStyle.WS_MAXIMIZEBOX)
            End If
         End If
         
         Const WM_NCACTIVATE = &H86&
         retVal = apiSendMessage(HWnd, WM_NCACTIVATE, True, 0)

      End With 'UDTMII
      
   End If

End Function


Public Function IsSysMenuItemEnabled( _
                                  ByVal HWnd As Long, _
                                  wID As eSysMenuID, _
                                  Optional retErrorSuccess As Long) As Boolean

   IsSysMenuItemEnabled = Not IsSysMenuItemDisabled(HWnd, wID, retErrorSuccess)

End Function

Public Function IsSysMenuItemDisabled( _
                                      ByVal HWnd As Long, _
                                      wID As eSysMenuID, _
                                      Optional retErrorSuccess As Long) As Boolean

   Dim retState As eMenuState

   retState = GetSysMenumItemState(HWnd, wID, retErrorSuccess)
   IsSysMenuItemDisabled = (retState = (retState Or eMenuState.MFS_GRAYED))

End Function

Public Function GetSysMenumItemState( _
                                     ByVal HWnd As Long, _
                                     Optional wID As eSysMenuID = eSysMenuID.SMSC_CLOSE, _
                                     Optional retErrorSuccess As Long) As eMenuState

   Dim utdMII As MenuItemInfo
   Dim hSysMenu As Long

   hSysMenu = apiGetSystemMenu(HWnd, 0)

   With utdMII
      .cbSize = Len(utdMII)
      .dwTypeData = VBA.String$(255, 0)
      .cch = Len(.dwTypeData)
      .fMask = eMENUITEMINFOMask.MIIM_STATE
      
      .wID = wID
      retErrorSuccess = apiGetMenuItemInfo(hSysMenu, .wID, False, utdMII)
      
      If retErrorSuccess Then
         'Debug.Print ".wID=" & .wID & "   " & eSysMenuIDDesc(wID) & "   retErrorSuccess=" & retErrorSuccess
         'Debug.Print ".fState=" & .fState & " " & eMenuStateDesc(.fState)
         GetSysMenumItemState = .fState
         Exit Function
      End If
      
      .wID = -wID
      retErrorSuccess = apiGetMenuItemInfo(hSysMenu, .wID, False, utdMII)
      'Debug.Print ".wID=" & -.wID & "   " & eSysMenuIDDesc(wID) & " r  etErrorSuccess=" & retErrorSuccess
      'Debug.Print ".fState=" & .fState & "   " & eMenuStateDesc(.fState)
      GetSysMenumItemState = .fState
   End With 'UTDMII

End Function

Public Function eMenuStateDesc( _
   Index As eMenuState) As String
   Dim retVal As String
   Select Case Index
   Case eMenuState.MFS_CHECKED  '= &H8&
      retVal = "MFS_CHECKED"
      'Checks the menu item. For more information about selected menu
      'items, see the hbmpChecked member.
   Case eMenuState.MFS_DEFAULT  '= &H1000&
      retVal = "MFS_DEFAULT"
      'Specifies that the menu item is the default. A menu can contain
      'only one default menu item, which is displayed in
      'bold.
   Case eMenuState.MFS_DISABLED  '= &H3&
      retVal = "MFS_DISABLED"
      'Disables the menu item and grays it so that it cannot be selected.
      'This is equivalent to MFS_GRAYED.
   Case eMenuState.MFS_ENABLED  '= &H0&
      retVal = "MFS_ENABLED"
      'Enables the menu item so that it can be selected. This is the
      'default state.
   Case eMenuState.MFS_GRAYED  '= &H3&
      retVal = "MFS_GRAYED"
      'Disables the menu item and grays it so that it cannot be selected.
      'This is equivalent to MFS_DISABLED.
   Case eMenuState.MFS_HILITE  '= &H80&
      retVal = "MFS_HILITE"
      'Highlights the menu item.
   Case eMenuState.MFS_UNCHECKED  '= &H0&
      retVal = "MFS_UNCHECKED"
      'Unchecks the menu item. For more information about clear menu
      'items, see the hbmpUnchecked member.
   Case eMenuState.MFS_UNHILITE  '= &H0&
      retVal = "MFS_UNHILITE"
      'Removes the highlight from the menu item. This is the default
      'state.
   Case Else
   End Select 'eMenuState
   eMenuStateDesc = retVal
End Function 'eMenuStateDesc

Public Function eSysMenuIDDesc( _
   Index As eSysMenuID) As String
   Dim retVal As String
   Select Case Index
      'this values are related to the WM_SYSCOMMAND message
      'and eWM_SYSCOMMAND eunumerator
   Case eSysMenuID.SMSC_CLOSE  '= &HF060&
      retVal = "SMSC_CLOSE"
   Case eSysMenuID.SMSC_DEFAULT  '= &HF160&
      retVal = "SMSC_DEFAULT"
   Case eSysMenuID.SMSC_MAXIMIZE  '= &HF030&
      retVal = "SMSC_MAXIMIZE"
   Case eSysMenuID.SMSC_MINIMIZE  '= &HF020&
      retVal = "SMSC_MINIMIZE"
   Case eSysMenuID.SMSC_MOVE  '= &HF010&
      retVal = "SMSC_MOVE"
   Case eSysMenuID.SMSC_SIZE  '= &HF000&
      retVal = "SMSC_SIZE"
   Case eSysMenuID.SMSC_RESTORE  '= &HF120&
      retVal = "SMSC_RESTORE"
   Case Else
   End Select 'eSysMenuID
   eSysMenuIDDesc = retVal
End Function 'eSysMenuIDDesc



