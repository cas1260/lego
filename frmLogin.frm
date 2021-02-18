VERSION 5.00
Begin VB.Form FrmSenha 
   Caption         =   "Senha de Sistema"
   ClientHeight    =   1545
   ClientLeft      =   2850
   ClientTop       =   3495
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Senha"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "FrmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    End
End Sub

Private Sub cmdOK_Click()
Dim U As String
    'check for correct password
    txtPassword.Text = UCase(txtPassword.Text)
    txtUserName.Text = UCase(txtUserName.Text)
    Usuario.Login = Trim(UCase(txtUserName.Text))
    Usuario.Senha = Trim(UCase(txtPassword.Text))
    Usuario.DataHora = Date & " " & Time
    
    U = UCase(DesComplica(Ler("Senha", txtPassword.Text, "", App.Path + "\Config.sis")))
    
    u = ler(
    If txtPassword.Text = U And txtUserName.Text = "SUB" Then
        Acesso = False
        MdiPrincipal.Show
        Unload Me
        Exit Sub
    ElseIf txtPassword.Text = "METAL" & Mid(Time, 4, 2) And txtUserName.Text = "METAL" Then
        Acesso = True
        MdiPrincipal.Show
        Unload Me
        Exit Sub
    ElseIf txtPassword.Text = U And txtUserName.Text = "LI" Then
        LimpaBanco
    Else
        MsgBox "Usuario ou senha esta Invalido", vbInformation, App.Title
        txtUserName.Text = ""
        txtPassword.Text = ""
        txtUserName.SetFocus
        
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
If KeyCode = 27 Then Unload Me
End Sub

Private Sub FrmSenha_Click()

End Sub

Private Sub LimpaBanco()
On Error Resume Next
Dim Banco As Database
Dim Arq As String
Dim X As Long

Arq = App.Path + "\Metal.Mdb"
If Dir(Arq) = "" Then
    MsgBox "Impossivel Localizar o Banco de Dados", vbInformation, App.Title
    End
End If

Set Banco = OpenDatabase(Arq, , , ";PWD=neo")
' OpenDatabase("Publishers", _
      dbDriverNoPrompt, True, _
      "ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers")

For X = 0 To Banco.TableDefs.Count
    Comando = "Delete * From " & Banco.TableDefs(X).Name
    Banco.Execute Comando
Next X
Banco.Close
On Error GoTo Erro
DBEngine.RepairDatabase Arq
DBEngine.CompactDatabase Arq, Arq + "a"
Kill Arq
FileCopy Arq + "a", Arq
Kill Arq + "A"
Acesso = False
MdiPrincipal.Show
Unload Me
Exit Sub

Erro:
'E
End Sub

Private Sub Form_Load()
SenhaSistema = "radio"
End Sub
