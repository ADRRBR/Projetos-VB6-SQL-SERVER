VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_DocClasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iStatus    As ADRRBR_SIS_Basico.eStatus
Private sMensagem  As String

Dim sVersao        As String
Dim sNome          As String
Dim clPropriedades As Collection
Dim clMetodos      As Collection
Dim clEventos      As Collection

Private Const sClasse = "SIS_DOCCLASSE"

Public Property Let Status(ByVal vNewValue As ADRRBR_SIS_Basico.eStatus)
    iStatus = vNewValue
End Property
Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Function Gera(pCaminhoClasse As String, pTreeView As Object, pExpandeTodosNodes As Boolean) As Boolean
    Dim oTreeView    As MSComctlLib.TreeView
    Dim sLinha       As String
    Dim sAux()       As String
    Dim sAux2()      As String
    Dim iAux         As Integer
    Dim sPropriedade As String
    Dim sMetodo      As String
    Dim sEvento      As String
    Dim sTipo        As String
    Dim sModo        As String
    Dim sRetorno     As String
    Dim sParametros  As String
    Dim bExiste      As Boolean
        
    On Error GoTo Erro_Gera

    LimpaStatus
    
    Gera = True
    iStatus = Sucesso
    
    Set oTreeView = pTreeView
    
    If Trim(pCaminhoClasse) = Empty Then
        Gera = False
        iStatus = Erro
        sMensagem = "Informe o nome do arquivo!"
        Exit Function
    End If
    
    If Dir(pCaminhoClasse) = Empty Then
        Gera = False
        iStatus = Erro
        sMensagem = "O arquivo: " & vbLf & vbLf & pCaminhoClasse & vbLf & vbLf & "Não existe!"
        Exit Function
    End If
    
    Open pCaminhoClasse For Input As #1
    
    If EOF(1) Then
        Gera = False
        iStatus = Erro
        sMensagem = "O arquivo: " & vbLf & vbLf & pCaminhoClasse & vbLf & vbLf & "Está vazio!"
        Exit Function
    End If
    
    sNome = Empty
    sVersao = Empty
    Set clPropriedades = New Collection
    Set clMetodos = New Collection
    Set clEventos = New Collection
    
    While Not EOF(1)
        Line Input #1, sLinha
        
        'Versão
        If UCase(Left(Trim(sLinha), 7)) = "VERSION" Then
            sAux() = Split(sLinha, " ")
            sVersao = "VERSAO;" & sAux(1)
        End If
    
        'Nome
        If UCase(Left(Trim(sLinha), 17)) = "ATTRIBUTE VB_NAME" Then
            sAux() = Split(sLinha, "=")
            sNome = "NOME;" & Trim(Replace(sAux(1), Chr(34), ""))
        End If
        
        'Propriedades
        If UCase(Left(Trim(sLinha), 15)) = "PUBLIC PROPERTY" Then
            If oBasico.Geral.PesquisaTexto(sLinha, "PUBLIC PROPERTY LET") Or oBasico.Geral.PesquisaTexto(sLinha, "PUBLIC PROPERTY SET") Then
                sModo = "Gravação"
                sAux() = Split(sLinha, " ")
                sPropriedade = sAux(3)
                sTipo = Trim(Replace(sAux(6), ")", ""))
                sAux() = Split(sPropriedade, "(")
                sPropriedade = sAux(0)
            
            ElseIf oBasico.Geral.PesquisaTexto(sLinha, "PUBLIC PROPERTY GET") Then
                sModo = "Leitura"
                sAux() = Split(sLinha, " ")
                sPropriedade = Trim(Replace(Replace(sAux(3), "(", ""), ")", ""))
                sTipo = Trim(Replace(sAux(5), ")", ""))
                sAux() = Split(sPropriedade, "(")
            End If
            
            bExiste = False
            For iAux = 1 To clPropriedades.Count
                sAux() = Split(clPropriedades(iAux), "#")
                sAux() = Split(sAux(0), ";")
                If sAux(1) = sPropriedade Then
                    sPropriedade = clPropriedades(iAux) & "-" & sTipo & "@" & sModo
                    clPropriedades.Remove iAux
                    clPropriedades.Add sPropriedade
                    bExiste = True
                    Exit For
                End If
            Next iAux
            If Not bExiste Then
                clPropriedades.Add "PROPRIEDADE;" & sPropriedade & "#DEFINICAO;" & sTipo & "@" & sModo
            End If
        End If
        
        'Métodos
        If UCase(Left(Trim(sLinha), 10)) = "PUBLIC SUB" Or UCase(Left(Trim(sLinha), 15)) = "PUBLIC FUNCTION" Then
            If oBasico.Geral.PesquisaTexto(sLinha, "PUBLIC SUB") Then
                sAux() = Split(sLinha, " ")
                sAux() = Split(sAux(2), "(")
                sMetodo = sAux(0)
                sRetorno = "Sem Retorno"
                
            ElseIf oBasico.Geral.PesquisaTexto(sLinha, "PUBLIC FUNCTION") Then
                sAux() = Split(sLinha, " ")
                sAux() = Split(sAux(2), "(")
                sMetodo = sAux(0)
                sAux() = Split(sLinha, ")")
                If Trim(sAux(1)) <> Empty Then
                    sAux() = Split(sAux(1), " ")
                    sRetorno = sAux(UBound(sAux()))
                Else
                    sRetorno = "Sem Retorno"
                End If
            End If
            
            If sRetorno = "Sem Retorno" And oBasico.Geral.PesquisaTexto(sLinha, "()") Then
                sMetodo = sMetodo & "    (Sem Retorno, Sem Parâmetros)"
                clMetodos.Add "METODO;" & sMetodo
                
            ElseIf sRetorno = "Sem Retorno" Or oBasico.Geral.PesquisaTexto(sLinha, "()") Then
                If sRetorno = "Sem Retorno" Then
                    GoSub LocalizaParametros
                    sMetodo = sMetodo & "    (Sem Retorno)"
                    clMetodos.Add "METODO;" & sMetodo & "#" & "#" & sParametros
                    
                ElseIf oBasico.Geral.PesquisaTexto(sLinha, "()") Then
                    sMetodo = sMetodo & "    (Sem Parâmetros)"
                    clMetodos.Add "METODO;" & sMetodo & "#" & "RETORNO;" & sRetorno & "#"
                End If
            Else
                GoSub LocalizaParametros
                clMetodos.Add "METODO;" & sMetodo & "#" & "RETORNO;" & sRetorno & "#" & sParametros
            End If
        End If
        
        'Eventos
        If UCase(Left(Trim(sLinha), 12)) = "PUBLIC EVENT" Then
            sAux() = Split(sLinha, " ")
            sAux() = Split(sAux(2), "(")
            sEvento = sAux(0)
            
            If oBasico.Geral.PesquisaTexto(sLinha, "()") Then
                sEvento = sEvento & "    (Sem Parâmetros)"
                clEventos.Add "EVENTO;" & sEvento
            Else
                GoSub LocalizaParametros
                clEventos.Add "EVENTO;" & sEvento & "#" & sParametros
            End If
        End If
    Wend
    
    Close #1
    
    iStatus = MontaNodesTreeviewClasse(pTreeView, pExpandeTodosNodes)
    
    If iStatus <> Sucesso Then
        Gera = False
        Set clPropriedades = Nothing
        Set clMetodos = Nothing
        Set clEventos = Nothing
        oBasico.Log.Grava sMensagem, "ERRO", sClasse, "MontaNodesTreeviewClasse"
        Exit Function
    End If
    
    Set clPropriedades = New Collection
    Set clMetodos = New Collection
    Set clEventos = New Collection
    
    Exit Function

LocalizaParametros:
    sAux() = Split(sLinha, "(")
    sAux() = Split(sAux(1), ")")
    sParametros = sAux(0)
    sAux() = Split(sParametros, ",")
    
    sParametros = "PARAMETROS;"
    
    For iAux = 0 To UBound(sAux())
        sAux2() = Split(sAux(iAux), " As ")
        
        sModo = Empty
        If oBasico.Geral.PesquisaTexto(sAux(iAux), "Optional") Then
            sModo = sModo & "Opcional"
            sAux2(0) = Replace(sAux2(0), "Optional", "")
        End If
        If oBasico.Geral.PesquisaTexto(sAux(iAux), "ByRef") Then
            If sModo <> Empty Then sModo = sModo & ", "
            sModo = sModo & "Por Refer."
            sAux2(0) = Replace(sAux2(0), "ByRef", "")
        End If
        If oBasico.Geral.PesquisaTexto(sAux(iAux), "ByVal") Then
            If sModo <> Empty Then sModo = sModo & ", "
            sModo = sModo & "Por Valor"
            sAux2(0) = Replace(sAux2(0), "ByVal", "")
        End If
        
        If UBound(sAux2()) = 0 Then
            sTipo = "    (Sem Tipo)"
        Else
            sTipo = "     (" & Trim(sAux2(UBound(sAux2()))) & ")"
        End If
        
        sParametros = sParametros & Trim(sAux2(0)) & "@" & sTipo
        If sModo <> Empty Then sParametros = sParametros & "    " & Trim(sModo)
        
        If iAux < UBound(sAux()) Then sParametros = sParametros & "-"
    Next iAux
Return

Erro_Gera:
    iStatus = Erro
    Gera = False
    Set clPropriedades = Nothing
    Set clMetodos = Nothing
    Set clEventos = Nothing
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O MÉTODO < cls" & sClasse & "." & "Gera > "
    oBasico.Log.Grava sMensagem, "ERRO", sClasse, "MontaNodesTreeviewClasse"
End Function

Private Function MontaNodesTreeviewClasse(pTreeView As Object, pExpandeTodosNodes As Boolean) As ADRRBR_SIS_Basico.eStatus
    Dim oTreeView      As MSComctlLib.TreeView
    Dim tvwNode        As Node
    
    Dim sChavePai      As String
    Dim sChave         As String
    Dim sConteudo      As String
    
    Dim sNivelClasse   As String
    Dim sNivelNome     As String
    Dim sNivelVersao   As String
    Dim sNivelPropr1   As String
    Dim sNivelPropr2   As String
    Dim sNivelMetodo1  As String
    Dim sNivelMetodo2  As String
    Dim sNivelEvento1  As String
    Dim sNivelEvento2  As String
    Dim sNivelModo     As String
    Dim sNivelParam    As String
    Dim sNivelRetorno  As String
        
    Dim iAux           As Integer
    Dim iAux2          As Integer
    Dim sAux()         As String
    Dim sAux2()        As String
     
    On Error GoTo ErroRotina
     
    MontaNodesTreeviewClasse = Sucesso
     
    Set oTreeView = pTreeView
    
    'Classe
    sNivelClasse = "Classe"
    sChavePai = Empty
    sChave = sNivelClasse
    sConteudo = "Classe"
    oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Azul, True, True
    GoSub VerificaStatus
    
    'Nome
    sAux() = Split(sNome, ";")
    sNivelNome = sAux(0)
    sChavePai = sNivelClasse
    sChave = sNivelNome
    sConteudo = sAux(1)
    oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Azul, True, True
    GoSub VerificaStatus

    'Versão
    sAux() = Split(sVersao, ";")
    sNivelVersao = sAux(0)
    sChavePai = sNivelNome
    sChave = sNivelVersao
    sConteudo = sAux(1)
    oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Azul, True, True
    GoSub VerificaStatus
    
    'Propriedades
    sNivelPropr1 = "PROPRIEDADES"
    sChavePai = sNivelVersao
    sChave = sNivelPropr1
    sConteudo = "Propriedades"
    If clPropriedades.Count = 0 Then sConteudo = sConteudo & "    (Sem Propriedades)"
    oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Azul, Preto, True
    GoSub VerificaStatus
    
    'Cada Propriedade
    For iAux = 1 To clPropriedades.Count
        'Propriedade
        sAux() = Split(clPropriedades(iAux), ";")
        sNivelPropr2 = sAux(0)
        sChavePai = sNivelPropr1
        sChave = sNivelPropr2 & iAux
        sAux() = Split(sAux(1), "#")
        sConteudo = sAux(0)
        oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, True, pExpandeTodosNodes
        GoSub VerificaStatus
        
        'Modo e Tipo da Propriedade
        sAux() = Split(clPropriedades(iAux), ";")
        sNivelModo = "MODO" & iAux
        sAux() = Split(sAux(2), "-")
        sChavePai = sNivelPropr2 & iAux
        For iAux2 = 0 To UBound(sAux())
            sChave = sNivelModo & (iAux2 + 1)
            sAux2() = Split(sAux(iAux2), "@")
            sConteudo = sAux2(1) & "    (" & sAux2(0) & ")"
            oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, False, pExpandeTodosNodes
            GoSub VerificaStatus
        Next iAux2
    Next iAux

    'Métodos
    sNivelMetodo1 = "METODOS"
    sChavePai = sNivelVersao
    sChave = sNivelMetodo1
    sConteudo = "Métodos"
    If clMetodos.Count = 0 Then sConteudo = sConteudo & "    (Sem Métodos)"
    oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Azul, Preto, True
    GoSub VerificaStatus

    'Cada Método
    For iAux = 1 To clMetodos.Count
        'Método
        sAux() = Split(clMetodos(iAux), ";")
        sNivelMetodo2 = sAux(0)
        sChavePai = sNivelMetodo1
        sChave = sNivelMetodo2 & iAux
        sAux() = Split(sAux(1), "#")
        sConteudo = sAux(0)
        oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Verde, True, pExpandeTodosNodes
        GoSub VerificaStatus
    
        'Retorno do Método
        If oBasico.Geral.PesquisaTexto(clMetodos(iAux), "RETORNO;") Then
            sAux() = Split(clMetodos(iAux), "#")
            sAux() = Split(sAux(1), ";")
            sNivelRetorno = sAux(0) & iAux
            sChavePai = sNivelMetodo2 & iAux
            sChave = sNivelRetorno
            sConteudo = "Retorno" & "    (" & sAux(1) & ")"
            oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, False, pExpandeTodosNodes
            GoSub VerificaStatus
        End If
        
        'Parâmetros do Método
        If oBasico.Geral.PesquisaTexto(clMetodos(iAux), "PARAMETROS;") Then
            sNivelParam = "PARAMETROSMETODO" & iAux & "."
            sChavePai = sNivelMetodo2 & iAux
            sChave = sNivelParam
            sConteudo = "Parâmetros"
            oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, True, pExpandeTodosNodes
            GoSub VerificaStatus
    
            'Cada Parâmetro do Método
            sAux() = Split(clMetodos(iAux), "#")
            sAux() = Split(sAux(2), ";")
            sAux() = Split(sAux(1), "-")
            sChavePai = sNivelParam
            For iAux2 = 0 To UBound(sAux())
                sChave = sNivelParam & (iAux2 + 1)
                sAux2() = Split(sAux(iAux2), "@")
                If UBound(sAux2()) > 0 Then
                    sConteudo = sAux2(0) & sAux2(1)
                Else
                    sConteudo = sAux2(0)
                End If
                oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, False, pExpandeTodosNodes
                GoSub VerificaStatus
            Next iAux2
        End If
    Next iAux
    
    'Eventos
    sNivelEvento1 = "EVENTOS"
    sChavePai = sNivelVersao
    sChave = sNivelEvento1
    sConteudo = "Eventos"
    If clEventos.Count = 0 Then sConteudo = sConteudo & "    (Sem Eventos)"
    oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Azul, Preto, True
    GoSub VerificaStatus

    'Cada Evento
    For iAux = 1 To clEventos.Count
        'Evento
        sAux() = Split(clEventos(iAux), ";")
        sNivelEvento2 = sAux(0)
        sChavePai = sNivelEvento1
        sChave = sNivelEvento2 & iAux
        sAux() = Split(sAux(1), "#")
        sConteudo = sAux(0)
        oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Vermelho, True, True
        GoSub VerificaStatus
    
        'Parâmetros do Evento
        sAux() = Split(clEventos(iAux), "#")
        If oBasico.Geral.PesquisaTexto(clEventos(iAux), "PARAMETROS;") Then
            sNivelParam = "PARAMETROSEVENTO" & iAux & "."
            sChavePai = sNivelEvento2 & iAux
            sChave = sNivelParam
            sConteudo = "Parâmetros"
            oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, True, pExpandeTodosNodes
            GoSub VerificaStatus
    
            'Cada Parâmetro do Evento
            sAux() = Split(clEventos(iAux), "#")
            sAux() = Split(sAux(1), ";")
            sAux() = Split(sAux(1), "-")
            sChavePai = sNivelParam
            For iAux2 = 0 To UBound(sAux())
                sChave = sNivelParam & (iAux2 + 1)
                sAux2() = Split(sAux(iAux2), "@")
                If UBound(sAux2()) > 0 Then
                    sConteudo = sAux2(0) & sAux2(1)
                Else
                    sConteudo = sAux2(0)
                End If
                oBasico.TV.AddNodeTV oTreeView, tvwNode, sChavePai, sChave, sConteudo, Preto, False, pExpandeTodosNodes
                GoSub VerificaStatus
            Next iAux2
        End If
    Next iAux

    Exit Function
    
VerificaStatus:
    If oBasico.TV.Status <> Sucesso Then
        iStatus = oBasico.TV.Status
        sMensagem = oBasico.TV.Mensagem
        sMensagem = sMensagem & " - " & sChavePai & "-" & sChave & "-" & sConteudo
        MontaNodesTreeviewClasse = iStatus
        Exit Function
    End If
Return
    
ErroRotina:
    MontaNodesTreeviewClasse = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR A ROTINA PRIVADA < cls" & sClasse & "." & "MontaNodesTreeviewClasse > "
    oBasico.Log.Grava sChavePai & "-" & sChave & "-" & sConteudo, sMensagem, sClasse, "MontaNodesTreeviewClasse"
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
