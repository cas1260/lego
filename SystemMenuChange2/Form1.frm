VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Enable/Disable System Buttons"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5085
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   2280
      Width           =   1995
   End
   Begin VB.CommandButton cmdGetEnableInfo 
      Caption         =   "Get Enable Info"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   60
      TabIndex        =   5
      Top             =   240
      Width           =   2595
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "Toggle Size Button"
      Height          =   315
      Left            =   2940
      TabIndex        =   4
      Top             =   1860
      Width           =   1995
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Toggle Move Button"
      Height          =   315
      Left            =   2940
      TabIndex        =   3
      Top             =   1440
      Width           =   1995
   End
   Begin VB.CommandButton cmdMax 
      Caption         =   "Toggle Max Button"
      Height          =   315
      Left            =   2940
      TabIndex        =   2
      Top             =   1020
      Width           =   1995
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "Toggle Min Button"
      Height          =   315
      Left            =   2940
      TabIndex        =   1
      Top             =   600
      Width           =   1995
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Toggle Close Button"
      Height          =   315
      Left            =   2940
      TabIndex        =   0
      Top             =   180
      Width           =   1995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_CLOSE
   cmdGetEnableInfo_Click
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdGetEnableInfo_Click()
   With Me.List1
      .Clear
      .AddItem "Close Enabled=" & IsSysMenuItemEnabled(Me.HWnd, SMSC_CLOSE)
      .AddItem "Maximize Enabled=" & IsSysMenuItemEnabled(Me.HWnd, SMSC_MAXIMIZE)
      .AddItem "Minimize Enabled=" & IsSysMenuItemEnabled(Me.HWnd, SMSC_MINIMIZE)
      .AddItem "Move Enabled=" & IsSysMenuItemEnabled(Me.HWnd, SMSC_MOVE)
      .AddItem "Restore Enabled=" & IsSysMenuItemEnabled(Me.HWnd, SMSC_RESTORE)
      .AddItem "Size Enabled=" & IsSysMenuItemEnabled(Me.HWnd, SMSC_SIZE)
   End With
End Sub

Private Sub cmdMax_Click()
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_MAXIMIZE
   cmdGetEnableInfo_Click
End Sub

Private Sub cmdMin_Click()
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_MINIMIZE
   cmdGetEnableInfo_Click
End Sub

Private Sub cmdMove_Click()
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_MOVE
   cmdGetEnableInfo_Click
End Sub

Private Sub cmdSize_Click()
   ToggleSysMenuEnableDisable Me.HWnd, SMSC_SIZE
   cmdGetEnableInfo_Click
End Sub

Private Sub Form_DblClick()
   Dim frm As Form1
   Set frm = New Form1
   frm.Show
End Sub
