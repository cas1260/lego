VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   4665
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8229
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7170
      Top             =   4830
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":0BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":14B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":2304
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":2BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":3E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "list.frx":4180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   675
      Left            =   7800
      TabIndex        =   0
      Top             =   4830
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X As Long
Dim itmX As ListItem

ListView1.ColumnHeaders.Add 1, "A", "Codigo"
ListView1.ColumnHeaders.Add 2, "B", "Nome"
ListView1.ColumnHeaders.Add 3, "c", "teste"
ListView1.ListItems.Clear
For X = 1 To 100
      Set itmX = ListView1.ListItems.Add(, , "Teste" + Str(X))
      itmX.SubItems(1) = "Cleber"
      itmX.SubItems(2) = "Chris"
          
    'ListView1.ListItems.Add 1, "A" + Str(X), X, 1
    'ListView1.ListItems.Add 2, "B" + Str(X), Chr(X + 100)
    'ListView1.ListItems.Add 1, "Aa" + Str(X), "Nome " + Str(X), 1
    'ListView1.ListItems.Add 2, "SS" + Str(X), "DDDDD" + Str(X), 1
    'ListView1.s
    'listview1.h
Next X
ListView1.StartLabelEdit
End Sub

Private Sub ListView1_Click()
'ListView1.StartLabelEdit
End Sub
