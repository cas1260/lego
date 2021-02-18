VERSION 5.00
Begin VB.Form FrmLixo 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2025
      Left            =   390
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FrmLixo.frx":0000
      Top             =   300
      Width           =   4485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   945
      Left            =   840
      TabIndex        =   0
      Top             =   2610
      Width           =   3645
   End
End
Attribute VB_Name = "FrmLixo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
   hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9

Private Sub Command1_Click()
   Dim curline As Long, numlines As Long
   
   ' Get line number from start of text selection
   curline = SendMessage(Text1.hwnd, EM_LINEFROMCHAR, Text1.SelStart, 0&)
   
   ' Get line count
   numlines = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0&, 0&)
   
   Caption = "Line " & (curline + 1) & " of " & numlines
End Sub

Private Sub Command11_Click()
Dim X As Long, Y As Long, Texto As String

For X = 0 To 99999
'    Y = InStr(1, Text1.Text, Chr(13))
'    If Y = 0 Then
'        Texto = Text1.Text
'    Else
'        Texto = Left(Text1.Text, Y + 1)
'    End If
'    Text1.Text = Right(Text1.Text, Len(Text1.Text) - Y - 1)
    'Texto = GetLine(Text1, X)

    Escreva "obj", "Codigo[" + Trim(Str(X)) + "]", Texto, "C:\Teste.Cas"
    Me.Caption = X
Next X
End Sub
