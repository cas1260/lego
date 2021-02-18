Attribute VB_Name = "Ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
    
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function Escreva(sSection As String, _
sKey As String, ByVal sValue As String, Arquivo As String) As Boolean
 
    Dim lR As Long
    Dim sItemValue As String
     
    sValue = Trim$(sValue)
    sItemValue = Trim$(sItemValue) & vbNullChar
    lR = WritePrivateProfileString(sSection, sKey, _
    sValue, Arquivo)

    If lR = 0 Then
        Escreva = False
    Else
        Escreva = True
    End If
End Function


Public Function Ler(sSection As String, _
       sKey As String, sDefault As String, Arquivo As String) As String
 
    Dim lR          As Long
    Dim sReturnedValue   As String
     
    sReturnedValue = Space$(512)
    lR = GetPrivateProfileString(sSection, sKey, sDefault, _
    sReturnedValue, 512, Arquivo)
    If lR = 0 Then
        Ler = vbNullString
    Else
        Ler = Left$(sReturnedValue, lR)
    End If
End Function


Sub Main()
    Dim sR As String
    Dim lR As String
    lR = Escreva("Windows", "Sistema", "Windows 98", "D:\Windows\Win.ini")
    sR = Ler("Windows", "Sistema", "", "D:\Windows\Win.Ini")
    MsgBox "INI Value: " & sR
End Sub



Public Function EscrevaVar(sSection As String, _
sKey As String, ByVal sValue As String) As Boolean
 
    Dim lR As Long
    Dim sItemValue As String, Arquivo As String
    
    Arquivo = "C:\TmpVar.win"
    sValue = sValue
    sItemValue = Trim$(sItemValue) & vbNullChar
    lR = WritePrivateProfileString(sSection, sKey, _
    sValue + ".", Arquivo)

    If lR = 0 Then
        EscrevaVar = False
    Else
        EscrevaVar = True
    End If
End Function


Public Function LerVar(sSection As String, _
       sKey As String, sDefault As String) As String
 
    Dim lR          As Long
    Dim sReturnedValue   As String, Arquivo As String
    
    Arquivo = "C:\TmpVar.win"
    sReturnedValue = Space$(512)
    lR = GetPrivateProfileString(sSection, sKey, sDefault, _
    sReturnedValue, 512, Arquivo)
    If lR = 0 Then
        LerVar = vbNullString
    Else
        LerVar = Left$(sReturnedValue, lR)
        If LerVar <> sDefault Then
            LerVar = Left(LerVar, Len(LerVar) - 1)
        End If
    End If
End Function


