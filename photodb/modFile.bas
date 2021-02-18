Attribute VB_Name = "modFile"
Option Explicit
Private Type OPENFILENAME
    lStructSize As Long
    HwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FileLengthags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
    Public Const OFN_READONLY = &H1
    Public Const OFN_OVERWRITEPROMPT = &H2
    Public Const OFN_HIDEREADONLY = &H4
    Public Const OFN_NOCHANGEDIR = &H8
    Public Const OFN_SHOWHELP = &H10
    Public Const OFN_ENABLEHOOK = &H20
    Public Const OFN_ENABLETEMPLATE = &H40
    Public Const OFN_ENABLETEMPLATEHANDLE = &H80
    Public Const OFN_NOVALIDATE = &H100
    Public Const OFN_ALLOWMULTISELECT = &H200
    Public Const OFN_EXTENSIONDIFFERENT = &H400
    Public Const OFN_PATHMUSTEXIST = &H800
    Public Const OFN_FILEMUSTEXIST = &H1000
    Public Const OFN_CREATEPROMPT = &H2000
    Public Const OFN_SHAREAWARE = &H4000
    Public Const OFN_NOREADONLYRETURN = &H8000
    Public Const OFN_NOTESTFILECREATE = &H10000
    Public Const OFN_NONETWORKBUTTON = &H20000
    Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
    Public Const OFN_EXPLORER = &H80000 ' new look commdlg
    Public Const OFN_NODEREFERENCELINKS = &H100000
    Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
    Public Const OFN_SHAREFALLTHROUGH = 2
    Public Const OFN_SHARENOWARN = 1
    Public Const OFN_SHAREWARN = 0

Public sHwnd As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Function SaveFile(Optional Filter As String, Optional Title As String, Optional InitDir As String) As String
    
    Dim OFN As OPENFILENAME
    Dim A As Long
    OFN.lStructSize = Len(OFN)
    OFN.HwndOwner = sHwnd
    OFN.hInstance = App.hInstance
    
    If Len(Filter) = 0 Then Filter = "Text Files (*.txt)" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    If Len(InitDir) = 0 Then InitDir = App.Path & "\"
    If Len(Title) = 0 Then Title = "Open file"
    
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"


    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    OFN.lpstrFilter = Filter
    OFN.lpstrFile = Space$(254)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = InitDir
    OFN.lpstrTitle = Title
    OFN.FileLengthags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    A = GetSaveFileName(OFN)


    If (A) Then
        SaveFile = Trim$(OFN.lpstrFile)
    Else
        SaveFile = ""
    End If
End Function


Public Function OpenFile(Optional Filter As String, Optional Title As String, Optional InitDir As String) As String
    
    Dim OFN As OPENFILENAME
    Dim A As Long
    Static LastDir As String
    If Len(LastDir) > 0 Then InitDir = LastDir
    If Len(Filter) = 0 Then Filter = "Text Files (*.txt)" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
    If Len(InitDir) = 0 Then InitDir = App.Path & "\"
    If Len(Title) = 0 Then Title = "Open file"
    
    OFN.lStructSize = Len(OFN)
    OFN.HwndOwner = sHwnd
    OFN.hInstance = App.hInstance
    
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    

    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    OFN.lpstrFilter = Filter
    OFN.lpstrFile = Space$(254)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = InitDir
    OFN.lpstrTitle = Title
    OFN.FileLengthags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    A = GetOpenFileName(OFN)


    If (A) Then
        OpenFile = Trim$(OFN.lpstrFile)
        LastDir = Left(OpenFile, Len(OpenFile) - InStr(1, StrReverse(OpenFile), "\"))
    Else
        OpenFile = ""
    End If
End Function



