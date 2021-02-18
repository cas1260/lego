VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6120
   ClientLeft      =   240
   ClientTop       =   1395
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   6120
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Lego 1.1"
   Begin VB.Timer T 
      Left            =   1248
      Top             =   5064
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   1848
      Top             =   3360
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   2565
      Picture         =   "frmSplash.frx":000C
      Top             =   5685
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Index           =   4
      Left            =   5730
      Picture         =   "frmSplash.frx":08D6
      Top             =   5670
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   720
      Index           =   3
      Left            =   5070
      Picture         =   "frmSplash.frx":29C8
      Top             =   1215
      Width           =   720
   End
   Begin VB.Image Image7 
      Height          =   480
      Index           =   2
      Left            =   4170
      Picture         =   "frmSplash.frx":3892
      Top             =   4935
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Index           =   1
      Left            =   2190
      Picture         =   "frmSplash.frx":558C
      Top             =   4755
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   4980
      Picture         =   "frmSplash.frx":63CE
      Top             =   4485
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   1695
      Picture         =   "frmSplash.frx":80C8
      Top             =   1470
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Index           =   0
      Left            =   2310
      Picture         =   "frmSplash.frx":8F0A
      Top             =   885
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5505
      Picture         =   "frmSplash.frx":9D4C
      Top             =   3540
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1620
      Picture         =   "frmSplash.frx":AB8E
      Top             =   3975
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3015
      Picture         =   "frmSplash.frx":B9D0
      Top             =   4980
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1410
      Picture         =   "frmSplash.frx":C812
      Top             =   2415
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5610
      Picture         =   "frmSplash.frx":E50C
      Top             =   2295
      Width           =   480
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Neo SoftWare"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   1416
      TabIndex        =   3
      Top             =   36
      Width           =   2352
   End
   Begin VB.Label lblProductName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lego 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2805
      TabIndex        =   2
      Top             =   150
      Width           =   3465
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cleber de Almeida Soares"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3696
      TabIndex        =   1
      Top             =   5880
      Width           =   1992
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   3000
      TabIndex        =   0
      Top             =   5880
      Width           =   828
   End
   Begin VB.Image imgLogo 
      Height          =   5772
      Left            =   0
      Picture         =   "frmSplash.frx":F34E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1356
   End
   Begin VB.Image Image1 
      Height          =   7260
      Left            =   885
      Picture         =   "frmSplash.frx":3651C
      Top             =   -720
      Width           =   5970
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    ReDim Run(100) As New FrmR
    ReDim Max(100) As Boolean, Min(100) As Boolean, Fecha(100) As Boolean
    ReDim TabRec(100) As TipoRecordSet
    ReDim BancoDeDados(100) As Database
    NomeSkin = "C:\Arquivos de programas\ActiveSkin 4\Skins\winaqua.skn"
    CRun.OpenExe = 1
    CRun.ComandoRun = Command '"C:\Meus documentos\Samples Lego\Demo.afs /Open"
    'CRun.ComandoRun = "C:\Windows\Desktop\WebAsp.afs /open"
    CRun.ComandoRun = Replace(UCase(CRun.ComandoRun), Chr(34), "")
  '  MsgBox CRun.ComandoRun
  '  End
    If Right(CRun.ComandoRun, 4) = "/RUN" Then
        CRun.OpenExe = 0
        Me.Visible = False
        CRun.ComandoRun = Left(CRun.ComandoRun, Len(CRun.ComandoRun) - 4)
        
        If Dir(CRun.ComandoRun) = "" Then
            MsgBox "Aplicativo não e valido Win32", vbCritical, App.Title
            End
        End If
        Timer1.Enabled = False
        Timer1.Interval = 0
        T.Interval = 1
        T.Enabled = True
        FrmPrincipal.Visible = False
        FrmPrincipal.LRun.Caption = "["
        FrmPrincipal.Show
    ElseIf Right(CRun.ComandoRun, 5) = "/OPEN" Then
        CRun.ComandoRun = Left(CRun.ComandoRun, Len(CRun.ComandoRun) - 5)
        CRun.OpenExe = 2
    End If
End Sub

Private Sub T_Timer()
Unload Me
End Sub

Private Sub Timer1_Timer()
FrmPrincipal.Show
Unload Me
End Sub

