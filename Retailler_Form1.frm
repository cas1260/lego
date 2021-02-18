VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   2430
   ClientTop       =   2295
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4980
   ScaleWidth      =   5985
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   330
      TabIndex        =   5
      Top             =   2010
      Width           =   2925
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3315
      TabIndex        =   4
      Top             =   930
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   225
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Interdire Redimension"
      Height          =   585
      Left            =   3300
      TabIndex        =   2
      Top             =   2505
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Autoriser Redimension"
      Height          =   495
      Left            =   165
      TabIndex        =   1
      Top             =   2505
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   1590
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   270
      Width           =   1725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Command1_Click()
    Retaillable Text1, True
    Retaillable List1, True
    Retaillable Command3, True
End Sub



Private Sub Command2_Click()
    Retaillable Text1, False
    Retaillable List1, False
    Retaillable Command3, False
End Sub

Private Sub DBGrid1_Click()

End Sub

Private Sub Form_Load()

'**************************************************
'* NOM : RedimCtl
'* DATE : 14/11/1998
'*
'* AUTEUR : Philippe Valar
'*
'* CODE TROUVE SUR "Le petit monde de Visual Basic"
'*                 http://www.vbasic.org
'*
'* DESCRIPTION :
'* Ce code permet de redimensionner sur la feuille
'* certains contrôles durant l'exécution du
'* programme.
'*
'**************************************************




End Sub


