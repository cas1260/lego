VERSION 5.00
Begin VB.Form FrmSenha 
   Caption         =   "Senha :"
   ClientHeight    =   1668
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4776
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1668
   ScaleWidth      =   4776
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   405
      Left            =   3150
      TabIndex        =   3
      Top             =   1200
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   405
      Left            =   1608
      TabIndex        =   2
      Top             =   1200
      Width           =   1545
   End
   Begin VB.TextBox TxtSenha 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   30
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   810
      Width           =   4605
   End
   Begin VB.Label Auto 
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Senha"
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If UCase(Senha) = UCase(TxtSenha.Text) Then
    SenhaOk = True
    Unload Me
Else
    SenhaOk = False
    MsgBox "Senha Invalida ! ! !", vbCritical, App.Title
    TxtSenha.Text = ""
    TxtSenha.SetFocus
End If
End Sub

Private Sub Command2_Click()
SenhaOk = False
If CRun.OpenExe = 0 Then
    End
Else
    Unload Me
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If SenhaOk = False And CRun.OpenExe = 0 Then End
End Sub

Private Sub TxtSenha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub
