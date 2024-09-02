Attribute VB_Name = "mdlSIS_Basico"
Option Explicit

'---------------------- TravaMaquina
Declare Function LockWorkStation Lib "user32.dll" () As Long
'----------------------

'---------------------- NomeComputador
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'----------------------

'---------------------- UsuarioLocal
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'----------------------

'---------------------- ConfereUsuarioLocal
Public Const LOGON32_LOGON_INTERACTIVE = 2
Public Const LOGON32_LOGON_BATCH = 4
Public Const LOGON32_LOGON_SERVICE = 5
Public Const LOGON32_PROVIDER_DEFAULT = 0

Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" _
(ByVal lpszUsername As String, ByVal lpszDomain As String, _
ByVal lpszPassword As String, ByVal dwLogonType As Long, _
ByVal dwLogonProvider As Long, phToken As Long) As Long

Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long
'----------------------

Public oBasico As New ADRRBR_SIS_Basico.clsSIS_Basico

Function CaracterAcentuado(pTexto As String) As Boolean
    Dim sCaracteres As String
    
    sCaracteres = "������������������������������������������������"
    
    If InStr(1, sCaracteres, pTexto) > 0 Then
        CaracterAcentuado = True
    Else
        CaracterAcentuado = False
    End If
End Function

Function RemoveAcentuacao(pTexto As String) As String
    RemoveAcentuacao = pTexto
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "a")
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "e")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "e")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "e")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "e")

    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "i")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "i")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "i")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "i")

    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "o")

    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "u")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "u")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "u")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "u")
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "C")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "c")
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "Y")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "�", "y")
End Function


