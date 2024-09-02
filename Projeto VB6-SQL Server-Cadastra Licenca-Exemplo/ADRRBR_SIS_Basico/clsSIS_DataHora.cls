VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_DataHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iStatus   As ADRRBR_SIS_Basico.eStatus
Private sMensagem As String

Private Const sClasse = "SIS_DataHora"

Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Function ValidarData(pDataSemFormatacao As String, pDataFormadata As String) As Boolean
    Dim iDia  As Integer
    Dim iMes  As Integer
    Dim lAno  As Long
    Dim sHora As String
    Dim sMsg  As String

    LimpaStatus

    ValidarData = True
    
    If pDataSemFormatacao <> Empty Then
        If Not DataValida(pDataFormadata, iDia, iMes, lAno, sHora, sMsg) Then
            iStatus = Erro
            sMensagem = "Informe uma data válida!"
            ValidarData = False
        End If
    End If
End Function

Public Function DataValida(pData As Variant, pDia As Integer, pMes As Integer, pAno As Long, pHora As String, Optional pMensagem As String) As Boolean
    Dim sAux()  As String
    Dim sAux2() As String
    Dim sAux3() As String
    Dim sAno    As String
    Dim sHora   As String
    
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    DataValida = True
    
    'Formato Correto (Hora/Minuto e ou Segundo)
    If InStr(1, pData, "/") = 0 Then
        GoSub InvalidarData
        Exit Function
    End If
    sAux = Split(pData, "/")
    If UBound(sAux()) > 2 Then
        GoSub InvalidarData
        Exit Function
    End If
    
    'Dia
    pDia = sAux(0)
    If pDia < 1 Or pDia > 31 Then
        GoSub InvalidarData
        Exit Function
    End If
    
    'Mes
    pMes = sAux(1)
    If pMes < 1 Or pMes > 12 Then
        GoSub InvalidarData
        Exit Function
    End If
    
    'Separa Ano e Hora
    sAux2 = Split(sAux(2), " ")
    
    'Ano
    If Len(Trim(sAux2(0))) < 4 Then
        GoSub InvalidarData
        Exit Function
    End If
    pAno = sAux2(0)
    
    'Hora
    If UBound(sAux2()) = 1 Then
        pHora = sAux2(1)
    End If
    
    'Separa Data e Hora
    sAux3 = Split(pData, " ")
    
    If Not IsDate(sAux3(0)) Then
        GoSub InvalidarData
        Exit Function
    End If
    
    Exit Function
    
InvalidarData:
    iStatus = Erro
    DataValida = False
    pMensagem = "A data informada < " & pData & " > é inválida!"
    sMensagem = pMensagem
    Exit Function
Return

ErroRotina:
    iStatus = Erro
    DataValida = False
    pMensagem = "Ao validar a data informada < " & pData & " >"
    pMensagem = pMensagem & ", ocorreu o erro: " & vbLf
    pMensagem = pMensagem & Err.Number & vbLf
    pMensagem = pMensagem & Err.Description & vbLf & vbLf
    pMensagem = pMensagem & "AO EXECUTAR O MÉTODO < " & sClasse & "." & "DataValida > "
    sMensagem = pMensagem
End Function

Public Function ValidarHora(pHoraSemFormatacao As String, pHoraFormatada As String) As Boolean
    ValidarHora = True
    
    LimpaStatus
    
    If pHoraSemFormatacao <> Empty Then
        If Not HoraValida(pHoraFormatada) Then
            iStatus = Erro
            sMensagem = "Informe uma hora válida!"
            ValidarHora = False
        End If
    End If
End Function

Public Function HoraValida(pHora As String, Optional pMensagem As String) As Boolean
    Dim sAux() As String
    
    On Error GoTo ErroRotina
    
    LimpaStatus
    
    HoraValida = True
    
    'Formato Correto (Hora/Minuto e ou Segundo)
    If InStr(1, pHora, ":") = 0 Then
        GoSub InvalidarHora
        Exit Function
    End If
    sAux = Split(pHora, ":")
    If UBound(sAux()) > 2 Then
        GoSub InvalidarHora
        Exit Function
    End If
    
    'Horas
    If Len(Trim(sAux(0))) < 2 Then
        GoSub InvalidarHora
        Exit Function
    End If
    If Val(sAux(0)) > 23 Then
        GoSub InvalidarHora
        Exit Function
    End If
    
    'Minutos
    If Len(Trim(sAux(1))) < 2 Then
        GoSub InvalidarHora
        Exit Function
    End If
    If sAux(1) > 59 Then
        GoSub InvalidarHora
        Exit Function
    End If
    
    'Segundos
    If UBound(sAux()) = 2 Then
        If Len(Trim(sAux(2))) < 2 Then
            GoSub InvalidarHora
            Exit Function
        End If
        If sAux(2) > 59 Then
            GoSub InvalidarHora
            Exit Function
        End If
    End If
    
    Exit Function
    
InvalidarHora:
    iStatus = Erro
    HoraValida = False
    pMensagem = "A hora informada < " & pHora & " > é inválida!"
    sMensagem = pMensagem
    Exit Function
Return

ErroRotina:
    iStatus = Erro
    HoraValida = False
    pMensagem = "Ao validar a hora informada < " & pHora & " >"
    pMensagem = pMensagem & ", ocorreu o erro: " & vbLf
    pMensagem = pMensagem & Err.Number & vbLf
    pMensagem = pMensagem & Err.Description & vbLf & vbLf
    pMensagem = pMensagem & "AO EXECUTAR O MÉTODO < " & sClasse & "." & "HoraValida > "
    sMensagem = pMensagem
End Function

Public Function ExibeData(pData As Date) As String
    Dim sDia As String * 2
    Dim sMes As String * 2
    Dim sAno As String * 4
    
    If Val(CStr(pData)) = 0 Then Exit Function
    
    sDia = Format(Day(pData), "00")
    sMes = Format(Month(pData), "00")
    sAno = Year(pData)

    ExibeData = sDia & "/" & sMes & "/" & sAno
End Function

Public Function DataInvertidaSemFormatacao(pData As Date) As String
    Dim sDia As String * 2
    Dim sMes As String * 2
    Dim sAno As String * 4
    
    If Val(CStr(pData)) = 0 Then Exit Function
    
    sDia = Format(Day(pData), "00")
    sMes = Format(Month(pData), "00")
    sAno = Year(pData)

    DataInvertidaSemFormatacao = sAno & sMes & sDia
End Function

Public Function FormataDataGravar(pDataFormatar As Variant) As String
    Dim iDia  As Integer
    Dim iMes  As Integer
    Dim lAno  As Long
    Dim sHora As String
    Dim sMsg  As String

    If Not DataValida(pDataFormatar, iDia, iMes, lAno, sHora, sMsg) Then
        FormataDataGravar = pDataFormatar
        Exit Function
    End If
    
    FormataDataGravar = Format(lAno, "0000") & "-" & Format(iMes, "00") & "-" & Format(iDia, "00") & " " & sHora
End Function

Public Function UpdateRecordsetColunaData(pData As String) As String
    'Esta Rotina Troca o Dia e o Mês de Uma Determinada Data no Formato ('MM/DD/AAAA HH:MM:SS')
    'Para Atualizar o Campo Data de um RecordSet na Memória!
    
    UpdateRecordsetColunaData = Mid(pData, 5, 2) & "/" & Mid(pData, 2, 2) & Mid(pData, 7, 14)
End Function

Private Sub LimpaStatus()
    On Error GoTo ErroRotina
    
    iStatus = SemRequisicao
    sMensagem = Empty
    
    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "LimpaStatus > "
End Sub

