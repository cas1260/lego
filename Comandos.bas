Attribute VB_Name = "Comandos"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public NomeVar() As String, ValorVar() As String, TipoVar() As String
Public FrmCodigoRum As Form
Public UltVar As Long, FimProg As Boolean
Public Nome_Da_Tela As String, Titulo_Pjt As String
Public Plataforma As String, Autor As String
Public Senha As String
Public NovoObj As Object
Public Vi As Visivel
Public ComEventos(9999) As EventosFinal
Public Max() As Boolean, Min() As Boolean, Fecha() As Boolean
Public SenhaOk As Boolean
Public NoOpen As Byte
Public FechaGeral As Boolean
Public TeclaKey As Long
Public TeclaEsc As Byte
Public TeclaEnter As Byte
Public TipoMenu As Boolean
Public LocalBancodeDados As String
Public Banco As Database
Public CRun As Exe
Public NomeRun As String
Public BancoDeDados() As Database
Public NomeSkin As String

Type Exe
    ComandoRun As String
    OpenExe As Long
End Type

Type Visivel
    Cmd(999) As String
    Lbl(999) As String
    Txt(999) As String
    Form(999) As String
    List(999) As String
    Cbo(999) As String
    Fm(999) As String
End Type
Type Totais
    TotalCmd As Long
    TotalTxt As Long
    TotalLbl As Long
    TotalFm As Long
    TotalImg As Long
    TotalOpt As Long
    TotalChk As Long
    TotalCombo As Long
    TotalLst As Long
    TotalBanco As Long
    TotalRecord As Long
    TotalTime As Long
End Type

Type Eventos_Obj
    Click As String
    DbClick As String
    GoFocus As String
    LostFocus As String
    KeyPress As String
    KeyDown As String
    Resize As String
    E_Load As String
    E_Close As String
End Type

Type EventosFinal
    Cmd As Eventos_Obj
    Form As Eventos_Obj
    Txt As Eventos_Obj
    Frame  As Eventos_Obj
    Lbl As Eventos_Obj
    Image As Eventos_Obj
    Combo As Eventos_Obj
    Opt As Eventos_Obj
    Check As Eventos_Obj
End Type
    
