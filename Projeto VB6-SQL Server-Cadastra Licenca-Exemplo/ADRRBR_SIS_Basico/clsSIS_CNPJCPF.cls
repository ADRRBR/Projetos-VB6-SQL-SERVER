VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_CNPJCPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum eTipoDocumento
    CPF = 1
    CNPJ = 2
End Enum

'***********************************
'      FUNÇÕES DE TERCEIROS
'***********************************
'--------------------------------------------------------
'Conjunto de funções para validação de CPF ou CNPJ
'--------------------------------------------------------
'----------------------------
'Calculo GERAL CPF CNPJ
'----------------------------
Public Function ValidaCPF_CNPJ(pTipoDocto As ADRRBR_SIS_Basico.eTipoDocumento, pNumero) As Boolean
    Select Case pTipoDocto
        Case CPF
            ValidaCPF_CNPJ = CalculaCPF(pNumero)
        Case CNPJ
            ValidaCPF_CNPJ = CalculaCNPJ(pNumero)
    End Select
End Function
'----------------------------
'fim Calculo GERAL CPF CNPJ
'----------------------------
'----------------------------
'Calcula CPF
'----------------------------
Private Function CalculaCPF(pCPF) As Boolean
    Dim S As String
    Dim ch As String
    Dim X As Integer

    Dim RecebeCPF, numero(11), soma, resultado1, resultado2
    RecebeCPF = pCPF
    'Retirar todos os caracteres que nao sejam 0-9

    S = ""
    For X = 1 To Len(RecebeCPF)
        ch = Mid(RecebeCPF, X, 1)
        If Asc(ch) >= 48 And Asc(ch) <= 57 Then
           S = S & ch
        End If
    Next
    RecebeCPF = S

    If Len(RecebeCPF) <> 11 Then
        CalculaCPF = False
    ElseIf RecebeCPF = "00000000000" Then
           CalculaCPF = False
        Else

            numero(1) = CInt(Mid(RecebeCPF, 1, 1))
            numero(2) = CInt(Mid(RecebeCPF, 2, 1))
            numero(3) = CInt(Mid(RecebeCPF, 3, 1))
            numero(4) = CInt(Mid(RecebeCPF, 4, 1))
            numero(5) = CInt(Mid(RecebeCPF, 5, 1))
            numero(6) = CInt(Mid(RecebeCPF, 6, 1))
            numero(7) = CInt(Mid(RecebeCPF, 7, 1))
            numero(8) = CInt(Mid(RecebeCPF, 8, 1))
            numero(9) = CInt(Mid(RecebeCPF, 9, 1))
            numero(10) = CInt(Mid(RecebeCPF, 10, 1))
            numero(11) = CInt(Mid(RecebeCPF, 11, 1))

            soma = 10 * numero(1) + 9 * numero(2) + 8 * numero(3) + 7 * numero(4) + 6 * numero(5) + 5 * numero(6) + 4 * numero(7) + 3 * numero(8) + 2 * numero(9)
            soma = soma - (11 * (Int(soma / 11)))

            If soma = 0 Or soma = 1 Then
                resultado1 = 0
            Else
                resultado1 = 11 - soma
            End If
    
            If resultado1 = numero(10) Then

            soma = numero(1) * 11 + numero(2) * 10 + numero(3) * 9 + numero(4) * 8 + numero(5) * 7 + numero(6) * 6 + numero(7) * 5 + numero(8) * 4 + numero(9) * 3 + numero(10) * 2
            soma = soma - (11 * (Int(soma / 11)))

            If soma = 0 Or soma = 1 Then
                resultado2 = 0
            Else
                resultado2 = 11 - soma
            End If

            If resultado2 = numero(11) Then
                CalculaCPF = True
            Else
                CalculaCPF = False
            End If
        Else
            CalculaCPF = False
        End If
    End If

End Function
'----------------------------
'CalculaCNPJ
'----------------------------
Private Function CalculaCNPJ(pCNPJ) As Boolean
    Dim S As String
    Dim ch As String
    Dim X As Integer

    Dim RecebeCNPJ, numero(14), soma, resultado1, resultado2
    RecebeCNPJ = pCNPJ

    S = ""
    For X = 1 To Len(RecebeCNPJ)
        ch = Mid(RecebeCNPJ, X, 1)
        If Asc(ch) >= 48 And Asc(ch) <= 57 Then
            S = S & ch
        End If
    Next
    RecebeCNPJ = S

    If Len(RecebeCNPJ) <> 14 Then
        CalculaCNPJ = False
    ElseIf RecebeCNPJ = "00000000000000" Then
        CalculaCNPJ = False
    Else

    numero(1) = CInt(Mid(RecebeCNPJ, 1, 1))
    numero(2) = CInt(Mid(RecebeCNPJ, 2, 1))
    numero(3) = CInt(Mid(RecebeCNPJ, 3, 1))
    numero(4) = CInt(Mid(RecebeCNPJ, 4, 1))
    numero(5) = CInt(Mid(RecebeCNPJ, 5, 1))
    numero(6) = CInt(Mid(RecebeCNPJ, 6, 1))
    numero(7) = CInt(Mid(RecebeCNPJ, 7, 1))
    numero(8) = CInt(Mid(RecebeCNPJ, 8, 1))
    numero(9) = CInt(Mid(RecebeCNPJ, 9, 1))
    numero(10) = CInt(Mid(RecebeCNPJ, 10, 1))
    numero(11) = CInt(Mid(RecebeCNPJ, 11, 1))
    numero(12) = CInt(Mid(RecebeCNPJ, 12, 1))
    numero(13) = CInt(Mid(RecebeCNPJ, 13, 1))
    numero(14) = CInt(Mid(RecebeCNPJ, 14, 1))

    soma = numero(1) * 5 + numero(2) * 4 + numero(3) * 3 + numero(4) * 2 + numero(5) * 9 + numero(6) * 8 + numero(7) * 7 + numero(8) * 6 + numero(9) * 5 + numero(10) * 4 + numero(11) * 3 + numero(12) * 2
    soma = soma - (11 * (Int(soma / 11)))

    If soma = 0 Or soma = 1 Then
        resultado1 = 0
    Else
        resultado1 = 11 - soma
    End If
    If resultado1 = numero(13) Then
        soma = numero(1) * 6 + numero(2) * 5 + numero(3) * 4 + numero(4) * 3 + numero(5) * 2 + numero(6) * 9 + numero(7) * 8 + numero(8) * 7 + numero(9) * 6 + numero(10) * 5 + numero(11) * 4 + numero(12) * 3 + numero(13) * 2
        soma = soma - (11 * (Int(soma / 11)))
        If soma = 0 Or soma = 1 Then
            resultado2 = 0
        Else
            resultado2 = 11 - soma
        End If
        If resultado2 = numero(14) Then
            CalculaCNPJ = True
        Else
            CalculaCNPJ = False
        End If
        Else
            CalculaCNPJ = False
        End If
    End If
End Function
'----------------------------
'fim CalculaCNPJ
'----------------------------
'--------------------------------------------------------
'fim Conjunto de funções para validação de CPF ou CNPJ
'--------------------------------------------------------
