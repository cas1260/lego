VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCompilar 
   Caption         =   "Lego"
   ClientHeight    =   5208
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8148
   Icon            =   "FrmCompilar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5208
   ScaleWidth      =   8148
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox F 
      BackColor       =   &H80000018&
      Height          =   3528
      Left            =   4296
      Pattern         =   "*.Exl"
      ReadOnly        =   0   'False
      System          =   -1  'True
      TabIndex        =   9
      Top             =   432
      Width           =   3804
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compliar"
      Height          =   348
      Left            =   6120
      TabIndex        =   7
      Top             =   4752
      Width           =   1908
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   348
      Left            =   4224
      TabIndex        =   8
      Top             =   4752
      Width           =   1908
   End
   Begin VB.TextBox TxtExe 
      ForeColor       =   &H8000000D&
      Height          =   288
      Left            =   72
      TabIndex        =   4
      Top             =   4296
      Width           =   7956
   End
   Begin VB.DirListBox Arq 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000002&
      Height          =   3528
      Left            =   72
      TabIndex        =   2
      Top             =   432
      Width           =   4212
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000018&
      Height          =   288
      Left            =   1056
      TabIndex        =   0
      Top             =   48
      Width           =   6996
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   252
      Left            =   48
      TabIndex        =   5
      Top             =   4824
      Visible         =   0   'False
      Width           =   4092
      _ExtentX        =   7218
      _ExtentY        =   445
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Label L 
      Caption         =   "Arquivo"
      ForeColor       =   &H8000000D&
      Height          =   192
      Left            =   72
      TabIndex        =   6
      Top             =   4632
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do Arquivo"
      ForeColor       =   &H8000000D&
      Height          =   204
      Left            =   96
      TabIndex        =   3
      Top             =   4032
      Width           =   1404
   End
   Begin VB.Label Label1 
      Caption         =   "Examinar &Em:"
      Height          =   204
      Left            =   48
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
End
Attribute VB_Name = "FrmCompilar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Arq_Change()
F.Path = Arq.Path
End Sub

Private Sub Command1_Click()
If Right(Arq.Path, 1) = "\" Then
    NomeRun = Arq.Path + TxtExe.Text
Else
    NomeRun = Arq.Path + "\" + TxtExe.Text
End If
If UCase(Right(NomeRun, 4)) <> ".EXL" Then
    NomeRun = NomeRun + ".Exl"
End If
If Dir(NomeRun) <> "" Then
    If MsgBox("O Arquivo já existe, Deseja Substituir ?" + Chr(13) + NomeRun, vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
        TxtExe.SetFocus
        Exit Sub
    End If
End If
FrmPrincipal.lp.Caption = "A"
Barra.Max = 500
Barra.Min = 0
Barra.Value = 0
Barra.Visible = True
Barra.Refresh
L.Caption = "Compilando..."
L.Refresh
Dim X As Long
For X = 0 To Barra.Max - 1
    Barra.Value = X
Next
Unload Me
        
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Arq.Path = Drive1.Drive
End Sub

Private Sub F_Click()
TxtExe.Text = F.FileName
End Sub

Private Sub Form_Load()
On Error Resume Next

End Sub

Private Sub TxtExe_GotFocus()
TxtExe.SelStart = 0
TxtExe.SelLength = Len(TxtExe.Text)
End Sub
