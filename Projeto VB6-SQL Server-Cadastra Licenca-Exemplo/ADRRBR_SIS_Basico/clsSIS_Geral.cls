VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Geral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lCor      As ADRRBR_SIS_Basico.eCor
Private iStatus   As ADRRBR_SIS_Basico.eStatus
Private sMensagem As String

Enum eCor
    Branco = &H80000005
    Preto = &H80000012
    CinzaBotao = &H8000000A
    Azul = &HC00000
    AzulSelecao = &H8000000D
    AzulGelo = &HE6E9E2
    Vermelho = &HC0&
    Verde = &H8000&
    VerdeEscuro = &H404000
    VerdeClaro = &H808000
    Laranja = &H80FF&
    Amarelo = &HC0FFFF
    AmareloEscuro = &HFFFF&
    Salmon = &HC0E0FF
    Marrom = &H404080
    Lilas = &H800080
    AbacateClaro = &HC0C0&
    Abacate = &H8080&
    AmareloClaro = &H80000018
    CorBotao = &H8000000F
End Enum

Enum eTeclas
    F1 = 112
    F2 = 113
    F3 = 114
    F4 = 115
    F5 = 116
    F6 = 117
    F7 = 118
    F8 = 119
    F9 = 120
    F10 = 121
    F11 = 122
    F12 = 123
    PGUp = 33
    PGDown = 34
End Enum

Enum eFormatoGravacaoRS
    PadraoRS = 0
    PadraoXML = 1
End Enum

Enum eAspas
    AspasSimples = 1
    AspasDuplas = 2
    CaracterSubstituicao = 3
End Enum

Enum eTipoAplicacao
    Windows = 1
    Internet = 2
End Enum

Private Const sClasse = "SIS_Geral"

Public Property Let Cor(ByVal vNewValue As ADRRBR_SIS_Basico.eCor)
    lCor = vNewValue
End Property
Public Property Get Cor() As ADRRBR_SIS_Basico.eCor
    Cor = lCor
End Property

Public Property Get TituloInclusao() As String
    TituloInclusao = "Inclusão"
End Property

Public Property Get TituloAlteracao() As String
    TituloAlteracao = "Alteração"
End Property

Public Property Get TituloConsulta() As String
    TituloConsulta = "Consulta"
End Property

Public Property Get TituloInclusaoConsulta() As String
    TituloInclusaoConsulta = "Inclusão/Consulta"
End Property

Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Sub PosicionaTela(pTela As Object, pFormulario As Object)
    pFormulario.Top = pTela.ActiveForm.Top + ((pTela.ActiveForm.Height - pFormulario.Height) / 2)
    pFormulario.Left = pTela.ActiveForm.Left + ((pTela.ActiveForm.Width - pFormulario.Width) / 2)
End Sub

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Function TrocaNuLL(Campo As Variant, Troca As Variant) As Variant
    TrocaNuLL = IIf(Trim(IsNull(Campo)), Troca, Campo)
End Function

Public Function SeleTexto(pControl As Object)
    'Seleciona todo o conteúdo do objeto
    If TypeOf pControl Is TextBox Then
        pControl.SelStart = 0
        pControl.SelLength = Len(pControl)
    End If
End Function

Public Function TrataNumeros(pAscii As Integer)
    'Excessões
    If pAscii = vbKeyBack Then
        TrataNumeros = pAscii
        Exit Function
    End If
    '***** Somente Números ******
    Const strNumeros = "0123456789"
    TrataNumeros = IIf(InStr(strNumeros, Chr(pAscii)) = 0, 0, pAscii)
End Function

Public Function Maiuscula(pAscii As Integer)
    Maiuscula = Asc(UCase(Chr(pAscii)))
End Function
    
Public Function Minuscula(pAscii As Integer)
    Minuscula = Asc(LCase(Chr(pAscii)))
End Function

Public Sub PosicionaCombo_Indice(pCombo As Object, pIndicePosicionar As Long)
    Dim oCombo  As Object
    Dim lIndice As Long
    
    Set oCombo = pCombo
    
    oCombo.ListIndex = -1
    
    For lIndice = 0 To oCombo.ListCount - 1
        If oCombo.ItemData(lIndice) = pIndicePosicionar Then
            oCombo.ListIndex = lIndice
            Exit For
        End If
    Next lIndice
End Sub

Public Sub PosicionaCombo_Conteudo(pCombo As Object, pConteudoPosicionar As String)
    Dim oCombo  As Object
    Dim lIndice As Long
    
    Set oCombo = pCombo
    
    oCombo.ListIndex = -1
    
    For lIndice = 0 To oCombo.ListCount - 1
        If oCombo.List(lIndice) = pConteudoPosicionar Then
            oCombo.ListIndex = lIndice
            Exit For
        End If
    Next lIndice
End Sub

Public Function MontaCalendario(pMSFlexGrid As Object, pMes As Integer, pAno As Long, Optional pDescricaoMes As String, Optional pFeriadosDelimVirgula As String, Optional pCorDomingo As ADRRBR_SIS_Basico.eCor, Optional pCorSabado As ADRRBR_SIS_Basico.eCor, Optional pCorFeriado As ADRRBR_SIS_Basico.eCor) As Boolean
    Dim oMSFlexGrid    As MSFlexGridLib.MSFlexGrid
    Dim rsDias         As ADODB.Recordset
    Dim dDataIni       As Date
    Dim dDataFim       As Date
    Dim iPrimDiaSemana As Integer
    Dim iDiaMes        As Integer
    Dim iDiasMes       As Integer
    Dim iDiaSem        As Integer
    Dim iLinha         As Integer
    Dim iColuna        As Integer
    Dim sFeriados()    As String
    Dim iFeriado       As Integer

    MontaCalendario = True
    
    Set oMSFlexGrid = pMSFlexGrid
    
    If (pMes < 1 Or pMes > 12) Or pAno = 0 Then
        MontaCalendario = False
        Exit Function
    End If
    
    pDescricaoMes = DescricaoMes(pMes)
    
    oBasico.FG.AddCol oMSFlexGrid, "Domingo", 800, Centro_FG, rsDias, PrimeiraCol_FG
    oBasico.FG.AddCol oMSFlexGrid, "Segunda", 800, Centro_FG, rsDias
    oBasico.FG.AddCol oMSFlexGrid, "Terça", 800, Centro_FG, rsDias
    oBasico.FG.AddCol oMSFlexGrid, "Quarta", 800, Centro_FG, rsDias
    oBasico.FG.AddCol oMSFlexGrid, "Quinta", 800, Centro_FG, rsDias
    oBasico.FG.AddCol oMSFlexGrid, "Sexta", 800, Centro_FG, rsDias
    oBasico.FG.AddCol oMSFlexGrid, "Sábado", 800, Centro_FG, rsDias, UltimaCol_FG
    
    dDataIni = CDate("01" & "/" & pMes & "/" & pAno)
    If pMes = 12 Then
        dDataFim = CDate("01" & "/" & "01" & "/" & CStr(pAno + 1)) - 1
    Else
        dDataFim = CDate("01" & "/" & CStr(pMes + 1) & "/" & pAno) - 1
    End If

    iPrimDiaSemana = Weekday(dDataIni)
    iDiasMes = DateDiff("D", dDataIni, dDataFim) + 1
    iDiaMes = 1
    
    While iDiaMes <= iDiasMes
        rsDias.AddNew
        
        For iDiaSem = iPrimDiaSemana To 7
            If iDiaMes <= iDiasMes Then
                rsDias.Fields(iDiaSem - 1).Value = iDiaMes
                iDiaMes = iDiaMes + 1
            End If
        Next iDiaSem
        
        iPrimDiaSemana = 1
    Wend
    
    oBasico.FG.AddReg oMSFlexGrid, rsDias, True
    
    'Domingos
    If pCorDomingo > 0 Then
        For iLinha = 1 To oMSFlexGrid.Rows - 1
            oMSFlexGrid.Row = iLinha: oMSFlexGrid.Col = 0
            oMSFlexGrid.CellForeColor = pCorDomingo
        Next iLinha
    End If
    
    'Sábados
    If pCorSabado > 0 Then
        For iLinha = 1 To oMSFlexGrid.Rows - 1
            oMSFlexGrid.Row = iLinha: oMSFlexGrid.Col = 6
            oMSFlexGrid.CellForeColor = pCorSabado
        Next iLinha
    End If
    
    'Feriados
    If pCorFeriado > 0 And InStr(1, pFeriadosDelimVirgula, ",") > 0 Then
        sFeriados = Split(pFeriadosDelimVirgula, ",")
        For iFeriado = 0 To UBound(sFeriados())
            For iLinha = 1 To oMSFlexGrid.Rows - 1
                For iColuna = 0 To oMSFlexGrid.Cols - 1
                    If sFeriados(iFeriado) <> Empty And oMSFlexGrid.TextMatrix(iLinha, iColuna) = sFeriados(iFeriado) Then
                        oMSFlexGrid.Row = iLinha: oMSFlexGrid.Col = iColuna
                        oMSFlexGrid.CellForeColor = pCorFeriado
                    End If
                Next iColuna
            Next iLinha
        Next iFeriado
    End If

    'Dia Atual
    For iLinha = 1 To oMSFlexGrid.Rows - 1
        For iColuna = 0 To oMSFlexGrid.Cols - 1
            If oMSFlexGrid.TextMatrix(iLinha, iColuna) = Day(Date) Then
                oMSFlexGrid.Row = iLinha: oMSFlexGrid.Col = iColuna
                oMSFlexGrid.CellBackColor = Vermelho: oMSFlexGrid.CellForeColor = Branco
                Exit For
            End If
        Next iColuna
    Next iLinha
End Function

Public Function DescricaoMes(pMes As Integer) As String
    If pMes < 1 Or pMes > 12 Then
        DescricaoMes = Empty
        Exit Function
    End If

    Select Case pMes
        Case 1: DescricaoMes = "Janeiro"
        Case 2: DescricaoMes = "Fevereiro"
        Case 3: DescricaoMes = "Março"
        Case 4: DescricaoMes = "Abril"
        Case 5: DescricaoMes = "Maio"
        Case 6: DescricaoMes = "Junho"
        Case 7: DescricaoMes = "Julho"
        Case 8: DescricaoMes = "Agosto"
        Case 9: DescricaoMes = "Setembro"
        Case 10: DescricaoMes = "Outubro"
        Case 11: DescricaoMes = "Novembro"
        Case 12: DescricaoMes = "Dezembro"
    End Select
End Function

Public Function GravaArquivoRS(pRSGravar As ADODB.Recordset, pCaminhoGravar As String, pFormatoGravar As eFormatoGravacaoRS) As Boolean
    On Error GoTo ErroGravaArquivoRS
    
    LimpaStatus
    
    GravaArquivoRS = True
    
    pRSGravar.Save pCaminhoGravar, pFormatoGravar
    
    Exit Function
    
ErroGravaArquivoRS:
    iStatus = Erro
    GravaArquivoRS = False
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "Não foi possivel gravar o arquivo < " & pCaminhoGravar & "> ."
End Function
    
Public Function LeArquivoPadraoRS(pCaminhoLer As String) As ADODB.Recordset
    On Error GoTo ErroLeArquivoPadraoRS
    
    LimpaStatus

    Set LeArquivoPadraoRS = Nothing
    Set LeArquivoPadraoRS = New ADODB.Recordset
    
    LeArquivoPadraoRS.Open pCaminhoLer
    
    Exit Function
    
ErroLeArquivoPadraoRS:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "Não foi possível ler o arquivo < " & pCaminhoLer & ">"
End Function

Public Function LeArquivoPadraoXML(pCaminhoLer As String) As ADODB.Recordset
    On Error GoTo ErroLeArquivoPadraoXML
    
    LimpaStatus
    
    GoTo ErroLeArquivoPadraoXML
    
    Exit Function
    
ErroLeArquivoPadraoXML:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "Não foi possível ler o arquivo < " & pCaminhoLer & ">"
End Function

Public Sub GravaRegistroWindows(pIdentificacao As String, pSubIdentificacao As String, pChave As String, pTextoGravar As String)
    On Error GoTo GravaRegistroWindows
    
    SaveSetting pIdentificacao, pSubIdentificacao, pChave, pTextoGravar
    
    Exit Sub
    
GravaRegistroWindows:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "Não foi possível gravar o registro do Windows!"
End Sub

Public Function LeRegistroWindows(pIdentificacao As String, pSubIdentificacao As String, pChave As String) As String
    On Error GoTo LeRegistroWindows
    
    LeRegistroWindows = GetSetting(pIdentificacao, pSubIdentificacao, pChave)
    
    Exit Function
    
LeRegistroWindows:
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "Não foi possível ler o registro do Windows!"
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

Public Function Aspas(pConteudo As String, pAspas As eAspas) As String
    Dim sCaracterSubstituicao As String
    
    On Error GoTo ErroAspas
    
    sCaracterSubstituicao = "^"
    
    Aspas = Replace(pConteudo, Chr(39), "{{ASS}}")
    Aspas = Replace(pConteudo, Chr(34), "{{ASD}}")
    
    Select Case pAspas
        Case AspasSimples: Aspas = Chr(39) & pConteudo & Chr(39)
        Case AspasDuplas: Aspas = Chr(34) & pConteudo & Chr(34)
        Case CaracterSubstituicao: Aspas = sCaracterSubstituicao & pConteudo & sCaracterSubstituicao
    End Select
    
    Aspas = Replace(Aspas, "{{ASS}}", Chr(39))
    Aspas = Replace(Aspas, "{{ASD}}", Chr(34))
    
    Exit Function
    
ErroAspas:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Aspas > "
End Function

Public Function EncriptarDecriptar(pTexto As String, pEncriptar As Boolean) As String
    Dim sTextoEncriptar As String
    Dim sCaracter       As String
    Dim iCaracter       As Integer
    Dim sTextoAux       As String
    
    sTextoAux = ":=:=:"
    
    For iCaracter = 1 To Len(Trim(pTexto))
        sCaracter = Mid(Trim(pTexto), iCaracter, 1)
        
        If pEncriptar Then
            If CaracterAcentuado(sCaracter) Then
                sTextoEncriptar = sTextoEncriptar & sTextoAux & sCaracter
            Else
                On Error GoTo ErroEncriptarMantemCaracter_1
                    sTextoEncriptar = sTextoEncriptar & Chr(Asc(sCaracter) + 125)
Continua_1:
                On Error GoTo 0
            End If
        Else
            If Mid(pTexto, iCaracter, Len(sTextoAux)) = sTextoAux Then
                iCaracter = iCaracter + Len(sTextoAux)
                sTextoEncriptar = sTextoEncriptar & Mid(pTexto, iCaracter, 1)
            Else
                sTextoEncriptar = sTextoEncriptar & Chr(Asc(sCaracter) - 125)
            End If
        End If
    Next
    
    EncriptarDecriptar = sTextoEncriptar
    
    Exit Function
    
ErroEncriptarMantemCaracter_1:
    sTextoEncriptar = sTextoEncriptar & sTextoAux & sCaracter
Resume Continua_1
End Function

Public Sub TravaMaquina()
    LockWorkStation
End Sub

Public Function NomeComputador() As String
    Dim Nome As String
    Dim Tamanho As Long
    
    Nome = Space(255)
    Tamanho = GetComputerName(Nome, 255)
    Nome = Trim(Replace(Nome, Chr$(0), " "))
    NomeComputador = Nome
End Function

Public Function UsuarioLocal() As String
    Dim sBuffer As String
    Dim lSize As Long

    sBuffer = Space$(255)
    lSize = Len(sBuffer)
    GetUserName sBuffer, lSize
    If lSize > 0 Then UsuarioLocal = Left$(sBuffer, lSize - 1)
End Function

Public Function ConfereUsuarioLocal(pUsuario As String, pDominio As String, pSenha As String) As Boolean
    Dim lStatus     As Long
    Dim TokenHandle As Long

    ConfereUsuarioLocal = False

    lStatus = LogonUser(pUsuario, pDominio, pSenha, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, TokenHandle)

    If lStatus Then
        ConfereUsuarioLocal = True
        lStatus = CloseHandle(TokenHandle)
    End If
End Function

Public Sub Pausa(pTempo As Long)
    'Informar o Parâmetro < pTempo > em Segundos
    
    Dim dTempoInicial As Date
    Dim dTempoAtual   As Date
    Dim lIntervalo    As Long
    
    dTempoInicial = Now()
                
    Do
        dTempoAtual = Now
        lIntervalo = DateDiff("s", dTempoInicial, dTempoAtual)
    Loop Until lIntervalo >= pTempo
End Sub

Public Function PesquisaTexto(pTextoBase As String, pTextoPesquisar As String, Optional pPosicaoEncontrada As Integer) As Boolean
    PesquisaTexto = False
    pPosicaoEncontrada = InStr(1, pTextoBase, pTextoPesquisar, vbTextCompare)
    If pPosicaoEncontrada > 0 Then PesquisaTexto = True
End Function

Public Sub FINALIZA_Forms()
    Dim iForm As Integer
    
    For iForm = Forms.Count - 1 To 0 Step -1
        On Error Resume Next
        Unload Forms(iForm)
        Set Forms = Nothing
        On Error GoTo 0
    Next iForm
End Sub
