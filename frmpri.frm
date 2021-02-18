VERSION 5.00
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00808000&
   Caption         =   "Lego 1.0"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmpri.frx":0000
   LinkTopic       =   "MDIForm1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Menu MenuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu menuNovo 
         Caption         =   "&Novo"
         Begin VB.Menu MenuProjeto 
            Caption         =   "Projeto"
         End
         Begin VB.Menu MenuBranco 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu MenuTela 
            Caption         =   "Tela"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu nadkf 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu MenuBranco2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSalvar 
         Caption         =   "&Salvar"
      End
      Begin VB.Menu MenuBranco4 
         Caption         =   "-"
      End
      Begin VB.Menu menuComplile 
         Caption         =   "&Complilar             "
         Shortcut        =   {F5}
      End
      Begin VB.Menu haskfd 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSair 
         Caption         =   "&Sair"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y As Boolean

Private Sub Form_Activate()
Y = True
End Sub

Private Sub Form_Load()
Me.Top = -60
Me.Left = -60
Me.Width = Screen.Width + 150
Me.Height = 690
FrmFerramentas.Show
FrmFerramentas.Height = FrmFerramentas.Height + 200
FrmPropreidade.Show
Y = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
    FrmFerramentas.Visible = True
End If
FrmFerramentas.WindowState = Me.WindowState
FrmPropreidade.WindowState = Me.WindowState
If Me.WindowState <> 1 Then
    If Y = True Then
        FrmFerramentas.Visible = True
    End If
End If

End Sub
