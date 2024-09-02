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
    
    sCaracteres = "ÃãÁáÀàÂâÄäÉéÈèÊêËëÍíÌìÎîÏïÕõÓóÒòÔôÖöÚúÙùÛûÜüÇçÝý"
    
    If InStr(1, sCaracteres, pTexto) > 0 Then
        CaracterAcentuado = True
    Else
        CaracterAcentuado = False
    End If
End Function

Function RemoveAcentuacao(pTexto As String) As String
    RemoveAcentuacao = pTexto
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ã", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ã", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Á", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "á", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "À", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "à", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Â", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "â", "a")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ä", "A")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ä", "a")
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "É", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "é", "e")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "È", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "è", "e")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ê", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ê", "e")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ë", "E")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ë", "e")

    RemoveAcentuacao = Replace(RemoveAcentuacao, "Í", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "í", "i")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ì", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ì", "i")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Î", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "î", "i")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ï", "I")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ï", "i")

    RemoveAcentuacao = Replace(RemoveAcentuacao, "Õ", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "õ", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ó", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ó", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ò", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ò", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ô", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ô", "o")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ö", "O")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ö", "o")

    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ú", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ú", "u")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ù", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ù", "u")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Û", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "û", "u")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ü", "U")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ü", "u")
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ç", "C")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ç", "c")
    
    RemoveAcentuacao = Replace(RemoveAcentuacao, "Ý", "Y")
    RemoveAcentuacao = Replace(RemoveAcentuacao, "ý", "y")
End Function


