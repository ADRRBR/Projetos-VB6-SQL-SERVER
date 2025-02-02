VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_TV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iStatus    As ADRRBR_SIS_Basico.eStatus
Private sMensagem  As String

Private Const sClasse = "SIS_TV"

Public Property Let Status(ByVal vNewValue As ADRRBR_SIS_Basico.eStatus)
    iStatus = vNewValue
End Property
Public Property Get Status() As ADRRBR_SIS_Basico.eStatus
    Status = iStatus
End Property

Public Property Get Mensagem() As String
    Mensagem = sMensagem
End Property

Public Sub AddNodeTV(pTreeView As Object, pNode As Object, pChavePai As String, pChave As String, pConteudo As String, pCor As eCor, pNegrito As Boolean, pExpande As Boolean, Optional pImageList As Object, Optional pIndImg As Integer)
    Dim oTreeView  As MSComctlLib.TreeView
    Dim oNode      As MSComctlLib.Node
    Dim oImageList As MSComctlLib.ImageList
    
    On Error GoTo ErroRotina
    
    LimpaStatus
    iStatus = Sucesso
    
    Set oTreeView = pTreeView
    Set oNode = pNode
    Set oImageList = pImageList
    
    If Trim(pChavePai) = Empty Then
        If oImageList Is Nothing Then
            oTreeView.Style = 6
            oTreeView.ImageList = Nothing
            Set oNode = oTreeView.Nodes.Add(, , pChave, pConteudo)
        Else
            oTreeView.Style = 7
            If oTreeView.ImageList Is Nothing Then oTreeView.ImageList = oImageList
            Set oNode = oTreeView.Nodes.Add(, , pChave, pConteudo, pIndImg)
        End If
    Else
        If oImageList Is Nothing Then
            Set oNode = oTreeView.Nodes.Add(pChavePai, tvwChild, pChave, pConteudo)
        Else
            Set oNode = oTreeView.Nodes.Add(pChavePai, tvwChild, pChave, pConteudo, pIndImg)
        End If
    End If
    
    oTreeView.Nodes(oTreeView.Nodes.Count).ForeColor = pCor
    oTreeView.Nodes(oTreeView.Nodes.Count).Bold = pNegrito
    oTreeView.Nodes(oTreeView.Nodes.Count).Expanded = pExpande

    Exit Sub

ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O M�TODO < cls" & sClasse & "." & "AddNodeTV > "
End Sub

Public Sub PosicionaTV(pTreeView As Object, Optional pChavePosicionar As String)
    Dim oTreeView  As MSComctlLib.TreeView
    Dim lIndice    As Long
    Dim bSelec     As Boolean
    
    On Error GoTo ErroRotina
    
    LimpaStatus
    iStatus = Sucesso
    
    Set oTreeView = pTreeView

    If oTreeView.Nodes.Count = 0 Then Exit Sub
    If oTreeView.Enabled = False Then Exit Sub
    
    bSelec = False
    For lIndice = 1 To oTreeView.Nodes.Count
        If oTreeView.Nodes(lIndice).Key = pChavePosicionar Then
            oTreeView.Nodes(lIndice).Selected = True
            oTreeView.Nodes(lIndice).EnsureVisible
            oTreeView.SetFocus
            bSelec = True
            Exit For
        End If
    Next lIndice

    If Not bSelec Then
        oTreeView.Nodes(1).Selected = True
        oTreeView.Nodes(1).EnsureVisible
        oTreeView.SetFocus
        Exit Sub
    End If

    Exit Sub
    
ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O M�TODO < cls" & sClasse & "." & "PosicionaTV > "
End Sub

Public Sub ControlaMarcacaoNiveisTV(pTreeView As Object, pNode As Object, pDelimitadorChaveConteudo As String, pDelimitadorNivel As String, Optional pInformaCheck As Boolean, Optional pCheck As Boolean)
    'Esta Rotina Controla a Marca��o/Desmarca��o (Check) dos N�veis de um Determinado
    'TreeView, Controlando Pais e Filhos a Partir do Elemento Marcado/Desmarcado.

    'Para Isso, � Necess�rio que Sejam Montadas as Chaves dos Elementos de Cada N�vel
    'Conforme o Exemplo:
    'Onde < DelimitadorChaveConteudo > = %
    '     < DelimitadorNivel > = ;
    
    'No Exemplo, o N�vel 1 � Representado Pela Letra < S > e
    '            o N�vel 2 � Representado Pela Letra < R >
    '            As Informa��es que Aparecem Depois do Caracterer Separador < % > �
    '            Particular da Informa��o em Cada Registro Em Cada Aplica��o, Como Por
    '            Exemplo o (C�digo do Cliente).
    
    '< % > - Separador de Tipo de N�vel com a Informa��o Particular do Registro.
    '        Ex.:  N�vel 1 -  S%15' , 'S%20'
    '              N�vel 2 -  R%14' , 'R%26', 'R%33'

    '< ; > - Separador de N�veis dos Elementos Relacionados
    '        Ex.: Primeiro Elemento do N�vel 2 Relacionado com o Primeiro Elemento do N�vel 1
    '             'S%15;R%14'
    '
    '             Segundo Elemento do N�vel 2 Relacionado com o Primeiro Elemento do N�vel 1
    '             'S%15;R%26'
    '
    '             Terceiro Elemento do N�vel 2 Relacionado com o Segundo Elemento do N�vel 1
    '             'S%20;R%33'
        
    Dim oTreeView            As MSComctlLib.TreeView
    Dim oNode                As MSComctlLib.Node
    Dim lNode                As Long
    Dim lNodePai             As Long
    Dim lFilhosMarcados      As Long
    Dim lFilhosDesmarcados   As Long
    Dim sNiveisNodeAtual()   As String
    Dim sNivelNodeAtual()    As String
    Dim sPrimeiroNivel       As String
    Dim sNivelIndicado       As String
    Dim sNivelCorrente       As String
    Dim sNodeAtual           As String
    Dim sNodePai             As String
    Dim sSeparaNiveisNode()  As String
    Dim lQtdeNiveisNodePai   As Long
    Dim lQtdeNiveisNodeAtual As Long

    On Error GoTo ErroRotina
    
    LimpaStatus
    iStatus = Sucesso
    
    Set oTreeView = pTreeView
    Set oNode = pNode

    'For�a a Marca��o/Desmarca��o das Informa��es no TreeView Caso o Par�metro
    '< pInformaCheck > For Verdadeiro.
    'Esta Op��o � Utilizada Para Carregar as Informa��es j� Gravados na Tabela, Exibido as
    'Marca��es no TreeView, Caso Contr�rio, a Marca��o/Desmarca��o � Controlada Pelo Pr�prio
    'Evento do TreeView de Onde Partiu a Chamada Desta Rotina.
    If pInformaCheck Then oNode.Checked = pCheck

    'Localiza o Primeiro N�vel para Controle, Pois, o Mesmo N�o Possui Pai!
    GoSub LocalizaPrimeiroNivel
    
    'Localiza o N�vel Indicado (Marcado/Desmarcado) Para Verificar se � ou n�o o Primeiro N�vel
    GoSub LocalizaNivelIndicado
    
    'Marca/Desmarca Pai Principal
    If sNivelIndicado <> sPrimeiroNivel Then 'Somente se N�o for o Primeiro N�vel...
        If oNode.Parent <> Empty Then 'Se Existir Pai...
            lNodePai = oNode.Index
            Do
                'Localiza o N�vel Corrente Para Verificar se � ou n�o o Primeiro N�vel
                GoSub LocalizaNivelCorrente
                
                If sNivelCorrente <> sPrimeiroNivel Then 'Somente se N�o for o Primeiro N�vel...
                    If oTreeView.Nodes(oTreeView.Nodes(lNodePai).Index).Parent <> Empty Then 'Se Existir Pai...
                        lNodePai = oTreeView.Nodes(oTreeView.Nodes(lNodePai).Index).Parent.Index
                        GoSub VerificaNivelModificado
                        oTreeView.Nodes(lNodePai).Expanded = True
                    Else
                        lNodePai = 0
                    End If
                Else
                    lNodePai = 0
                End If
            Loop Until lNodePai = 0
        End If
    End If

    GoSub MarcaDesmarcaFilhosRelacionados

    Exit Sub

LocalizaPrimeiroNivel:
    sNodeAtual = oTreeView.Nodes(oTreeView.Nodes(1).Index).Key
    
    sNiveisNodeAtual() = Split(sNodeAtual, pDelimitadorNivel)
    sNivelNodeAtual() = Split(sNiveisNodeAtual(UBound(sNiveisNodeAtual())), pDelimitadorChaveConteudo) 'O �ltimo Elemento � o N�vel Atual
    
    'O Primeiro Elemento � o Tipo de N�vel
    sPrimeiroNivel = sNivelNodeAtual(0) 'O Primeiro N�vel N�o Possui Pai!
Return
    
LocalizaNivelIndicado:
    sNodeAtual = oNode.Key
    
    sNiveisNodeAtual() = Split(sNodeAtual, pDelimitadorNivel)
    sNivelNodeAtual() = Split(sNiveisNodeAtual(UBound(sNiveisNodeAtual())), pDelimitadorChaveConteudo) 'O �ltimo Elemento � o N�vel Atual
    
    'O Primeiro Elemento � o Tipo de N�vel
    sNivelIndicado = sNivelNodeAtual(0) 'O Primeiro N�vel N�o Possui Pai!
Return
    
LocalizaNivelCorrente:
    sNodeAtual = oTreeView.Nodes(lNodePai).Key
    
    sNiveisNodeAtual() = Split(sNodeAtual, pDelimitadorNivel)
    sNivelNodeAtual() = Split(sNiveisNodeAtual(UBound(sNiveisNodeAtual())), pDelimitadorChaveConteudo) 'O �ltimo Elemento � o N�vel Atual
    
    'O Primeiro Elemento � o Tipo de N�vel
    sNivelCorrente = sNivelNodeAtual(0) 'O Primeiro N�vel N�o Possui Pai!
Return
    
VerificaNivelModificado:
    'Verifica se Existem Outros Filhos Marcados/Desmarcados no Mesmo N�vel do N�vel Marcado/Desmarcado Atualmente!
    lFilhosMarcados = 0
    lFilhosDesmarcados = 0
    
    sNodePai = oTreeView.Nodes(oTreeView.Nodes(lNodePai).Index).Key

    For lNode = oTreeView.Nodes.Count To lNodePai Step -1
        sNodeAtual = oTreeView.Nodes(oTreeView.Nodes(lNode).Index).Key

        'O Primeiro Elemento � o Tipo de N�vel
        If sNodePai <> sNodeAtual Then 'Se o Node Pesquisado For do Mesmo N�vel do Node Marcado/Desmarcado, Considera
            If InStr(1, sNodeAtual, sNodePai) Then
                sSeparaNiveisNode() = Split(sNodePai, pDelimitadorNivel)
                lQtdeNiveisNodePai = UBound(sSeparaNiveisNode()) + 1
                
                sSeparaNiveisNode() = Split(sNodeAtual, pDelimitadorNivel)
                lQtdeNiveisNodeAtual = UBound(sSeparaNiveisNode()) + 1
            
                If lQtdeNiveisNodeAtual = lQtdeNiveisNodePai + 1 Then 'Se o Node Pesquisado For do Pr�ximo N�vel com Rela��o ao Pai Atual, Verifica se Existem Filhos Marcados/Desmarcados...
                    If oTreeView.Nodes(lNode).Checked = True Then
                        lFilhosMarcados = lFilhosMarcados + 1
                    Else
                        lFilhosDesmarcados = lFilhosDesmarcados + 1
                    End If
                End If
            End If
        End If
    Next lNode

    'Marca/Desmarca o Pai...
    If oNode.Checked = False Then
        If lFilhosMarcados = 0 Then 'Se N�o Existem Filhos Marcados no Pr�ximo N�vel com Rela��o ao Pai Atual...
            oTreeView.Nodes(lNodePai).Checked = oNode.Checked 'Marca/Desmarca o Pai!
        End If
    Else
        oTreeView.Nodes(lNodePai).Checked = oNode.Checked 'Marca/Desmarca o Pai!
    End If
Return

MarcaDesmarcaFilhosRelacionados:
    'Marca/Desmarca os Filhos Relacionados ao Pai Marcado/Desmarcado
    For lNode = oNode.Index To oTreeView.Nodes.Count
        sNodeAtual = oTreeView.Nodes(oTreeView.Nodes(lNode).Index).Key

        If InStr(1, sNodeAtual, oNode.Key) Then
            oTreeView.Nodes(lNode).Checked = oNode.Checked 'Marca/Desmarca o Pai!
        End If
    Next lNode
Return

ErroRotina:
    iStatus = Erro
    sMensagem = "Ocorreu o erro: " & vbLf
    sMensagem = sMensagem & Err.Number & vbLf
    sMensagem = sMensagem & Err.Description & vbLf & vbLf
    sMensagem = sMensagem & "AO EXECUTAR O M�TODO < cls" & sClasse & "." & "ControlaMarcacaoNiveisTV > "
End Sub

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

