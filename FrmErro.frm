VERSION 5.00
Begin VB.Form FrmErro 
   Caption         =   "Erro"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5832
   ControlBox      =   0   'False
   Icon            =   "FrmErro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5832
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   405
      Left            =   3960
      TabIndex        =   1
      Top             =   1590
      Width           =   1635
   End
   Begin VB.Label Erro 
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "FrmErro.frx":0442
      Stretch         =   -1  'True
      Top             =   150
      Width           =   600
   End
End
Attribute VB_Name = "FrmErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X As Long
FimProg = False
Unload Me
For X = 0 To 100
    Unload Run(X)
Next X
End Sub

