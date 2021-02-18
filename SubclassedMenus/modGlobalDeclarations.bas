Attribute VB_Name = "modGlobalDeclarations"
Option Explicit

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'Record previous window handler
Public ProcOld As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Const TPM_CENTERALIGN = &H4
Public Const TPM_RIGHTALIGN = &H8
Public Const TPM_BOTTOMALIGN = &H20
Public Const TPM_VCENTERALIGN = &H10
Public Const TPM_TOPALIGN = &H0&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_RETURNCMD = &H100&

Public Const MF_CHECKED = &H8&
Public Const MF_UNCHECKED = &H0&
Public Const MF_APPEND = &H100&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_HELP = &H4000&
Public Const MFS_DEFAULT = &H1000&
Public Const MFS_ENABLED = &H1000&

Public Const MFT_OWNERDRAW = &H100&

Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10
Public Const MIIM_DATA = &H20

Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const mnuFileAnmlDog = 1
Public Const mnuFileAnmlCat = 2
Public Const mnuFileAnmlHamster = 3
Public Const mnuFileAnmlChinchilla = 4
Public Const mnuFileExit = 5
Public Const mnuFilePrchCash = 6
Public Const mnuFilePrchCredit = 7
Public Const mnuHelpAbout = 8


