VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_FG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lRow       As Long
Private lCol       As Long
Private lUltRow    As Long
Private lUltCol    As Long

Enum eAlinhamentoFG
    Esquerda_FG = 0
    Centro_FG = 3
    Direita_FG = 6
End Enum

Enum ePosicaoColunaFG
    PrimeiraCol_FG = 1
    UltimaCol_FG = 2
End Enum

Public Property Let Row(ByVal vNewValue As Long)
    lRow = vNewValue
End Property
Public Property Get Row() As Long
    Row = lRow
End Property

Public Property Let Col(ByVal vNewValue As Long)
    lCol = vNewValue
End Property
Public Property Get Col() As Long
    Col = lCol
End Property

Public Property Let UltRow(ByVal vNewValue As Long)
    lUltRow = vNewValue
End Property
Public Property Get UltRow() As Long
    UltRow = lUltRow
End Property

Public Property Let UltCol(ByVal vNewValue As Long)
    lUltCol = vNewValue
End Property
Public Property Get UltCol() As Long
    UltCol = lUltCol
End Property

'----------------------------------------
'1 - < AddCol > Cria Colunas no MSFLEXGRID, Criando Simultaneamente as Colunas no RecordSet
'2 - Atualiza o RecordSet, Conforme a Rotina do Programador
'3 - < AddReg > Atualiza o MSFLEXGRID
'ou
'1 - Atualiza o RecordSet, Conforme a Rotina do Programador
'2 - < AddCol > Cria Colunas no MSFLEXGRID, Criando Simultaneamente as Colunas no RecordSet
'3 - < AddReg > Atualiza o MSFLEXGRID
'----------------------------------------
Public Sub AddCol(pMSFlexGrid As Object, pColuna As String, pTamanho As Long, pAlinhamento As ADRRBR_SIS_Basico.eAlinhamentoFG, pRSFlexGrid As ADODB.Recordset, Optional pStatusColuna As ADRRBR_SIS_Basico.ePosicaoColunaFG)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    
    Set oMSFlexGrid = pMSFlexGrid
    
    With oMSFlexGrid
        If pStatusColuna = PrimeiraCol_FG Then
            .FixedRows = 1
            .Rows = 1
            .FixedCols = 0
            .Cols = 0
        End If
        .Cols = .Cols + 1
        .Row = 0
        .Col = .Cols - 1
        
        .ColWidth(.Col) = pTamanho
        .ColAlignment(.Col) = pAlinhamento
        .TextMatrix(.Row, .Col) = pColuna
    End With
    
    If pStatusColuna = PrimeiraCol_FG Then
        If pRSFlexGrid Is Nothing Then
            Set pRSFlexGrid = New ADODB.Recordset
        End If
        lUltRow = 0
        lUltCol = 0
    End If
    
    If pRSFlexGrid.State = 0 Then
        pRSFlexGrid.Fields.Append pColuna, adVarChar, "1000"
    End If
    
    If pStatusColuna = UltimaCol_FG Then
        If pRSFlexGrid.State = 0 Then
            pRSFlexGrid.Open
        End If
    End If
End Sub

Public Sub AddReg(pMSFlexGrid As Object, pRSFlexGrid As ADODB.Recordset, pInicializarGrid As Boolean)
    Dim oMSFlexGrid     As MSFlexGridLib.MSFlexGrid
    Dim lRow            As Long
    Dim lCol            As Long
    Dim bAdicionaColuna As Boolean
    
    Set oMSFlexGrid = pMSFlexGrid
    
    If pRSFlexGrid.RecordCount = 0 Then Exit Sub
    
    If pInicializarGrid Then
        oMSFlexGrid.Rows = 2: oMSFlexGrid.Cols = 2
        oMSFlexGrid.Clear
    End If
    
    If oMSFlexGrid.Cols = 0 Then bAdicionaColuna = True
    
    If Not bAdicionaColuna Then
        'Verifica se é a Configuração Padrão do MSFlexGrid, Quando Adicionada à Aplicação
        If oMSFlexGrid.Rows = 2 And oMSFlexGrid.Cols = 2 And oMSFlexGrid.TextMatrix(0, 0) = Empty And oMSFlexGrid.TextMatrix(0, 1) = Empty Then
            bAdicionaColuna = True
        End If
    End If
    
    If bAdicionaColuna Then
        For lCol = 0 To pRSFlexGrid.Fields.Count - 1
            Select Case lCol
                Case 0
                    oBasico.FG.AddCol oMSFlexGrid, pRSFlexGrid.Fields(lCol).Name, 1000, Esquerda_FG, pRSFlexGrid, PrimeiraCol_FG
                
                Case pRSFlexGrid.Fields.Count - 1
                    oBasico.FG.AddCol oMSFlexGrid, pRSFlexGrid.Fields(lCol).Name, 1000, Esquerda_FG, pRSFlexGrid, UltimaCol_FG
                
                Case Else
                    oBasico.FG.AddCol oMSFlexGrid, pRSFlexGrid.Fields(lCol).Name, 1000, Esquerda_FG, pRSFlexGrid
            End Select
        Next lCol
    End If
    
    pRSFlexGrid.MoveFirst
    
    While Not pRSFlexGrid.EOF
        oMSFlexGrid.Rows = oMSFlexGrid.Rows + 1
        lRow = oMSFlexGrid.Rows - 1
        
        For lCol = 0 To pRSFlexGrid.Fields.Count - 1
            If pRSFlexGrid.Fields(lCol).Type = adCurrency Then
                oMSFlexGrid.TextMatrix(lRow, lCol) = Format(oBasico.Geral.TrocaNuLL(pRSFlexGrid.Fields(lCol).Value, "0"), "###,###,###,##0.00")
            Else
                oMSFlexGrid.TextMatrix(lRow, lCol) = oBasico.Geral.TrocaNuLL(pRSFlexGrid.Fields(lCol).Value, Empty)
            End If
        Next lCol
        
        pRSFlexGrid.MoveNext
    Wend
End Sub

Public Sub SelecionaCelula(pMSFlexGrid As Object, Optional pCorLinha As ADRRBR_SIS_Basico.eCor, Optional pCorFundo As ADRRBR_SIS_Basico.eCor, Optional pSelecionaColunaVazia As Boolean)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    
    Set oMSFlexGrid = pMSFlexGrid
    
    lRow = oMSFlexGrid.Row
    lCol = oMSFlexGrid.Col

    If Not pSelecionaColunaVazia Then
        If oMSFlexGrid.TextMatrix(lRow, lCol) = Empty Then
            lRow = lUltRow
            lCol = lUltCol
        End If
    End If
    
    If lUltRow > 0 Then
        oMSFlexGrid.Row = lUltRow
        oMSFlexGrid.Col = lUltCol
        
        If pCorLinha > 0 Then
            oMSFlexGrid.CellForeColor = pCorLinha
        Else
            oMSFlexGrid.CellForeColor = Preto
        End If
        
        If pCorFundo > 0 Then
            oMSFlexGrid.CellBackColor = pCorFundo
        Else
            oMSFlexGrid.CellBackColor = Branco
        End If
    End If

    oMSFlexGrid.Row = lRow
    oMSFlexGrid.Col = lCol
    oMSFlexGrid.CellForeColor = oMSFlexGrid.CellBackColor
    oMSFlexGrid.CellBackColor = AzulSelecao
    
    lUltRow = lRow
    lUltCol = lCol
End Sub

Public Sub ManutencaoCelulaAtual(pMSFlexGrid As Object, pCampoManut As Object, pSelecionaTexto As Boolean)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    Dim sConteudo   As String
    Dim iDia        As Integer
    Dim iMes        As Integer
    Dim lAno        As Long
    Dim sHora       As String
    Dim sMsg        As String
    
    Set oMSFlexGrid = pMSFlexGrid
    If oMSFlexGrid.Row = 0 Then Exit Sub
    
    sConteudo = oMSFlexGrid.TextMatrix(oMSFlexGrid.Row, oMSFlexGrid.Col)
    
    pCampoManut.Font = oMSFlexGrid.Font
    pCampoManut.Top = oMSFlexGrid.Top + oMSFlexGrid.CellTop
    pCampoManut.Left = oMSFlexGrid.Left + oMSFlexGrid.CellLeft
    pCampoManut.Width = oMSFlexGrid.CellWidth + 5
    
    Select Case UCase(TypeName(pCampoManut))
        Case "TEXTBOX"
            pCampoManut.FontSize = oMSFlexGrid.FontSize
            pCampoManut.Height = oMSFlexGrid.CellHeight
            pCampoManut.Text = sConteudo
            pCampoManut.SelStart = 0
            If pSelecionaTexto Then oBasico.Geral.SeleTexto pCampoManut
        
        Case "COMBOBOX"
            pCampoManut.FontSize = 4
            oBasico.Geral.PosicionaCombo_Conteudo pCampoManut, sConteudo
            
        Case "MASKEDBOX"
            pCampoManut.FontSize = oMSFlexGrid.FontSize
            pCampoManut.Height = oMSFlexGrid.CellHeight
            pCampoManut.Mask = Empty
            pCampoManut.Text = Empty
            If oBasico.Geral.TrocaNuLL(sConteudo, Empty) <> Empty Then
                If oBasico.DataHora.DataValida(sConteudo, iDia, iMes, lAno, sHora, sMsg) Then
                    pCampoManut.Text = Format(sConteudo, "dd/mm/yyyy")
                End If
            End If
            pCampoManut.Mask = "##/##/####"
            pCampoManut.SelStart = 0
            If pSelecionaTexto Then oBasico.Geral.SeleTexto pCampoManut
    End Select
    
    pCampoManut.ZOrder 0
    pCampoManut.Visible = True
    pCampoManut.SetFocus
End Sub

Public Sub AtualizaCelulaAtual(pMSFlexGrid As Object, pCampoManut As Object)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    
    Set oMSFlexGrid = pMSFlexGrid

    oMSFlexGrid.TextMatrix(oMSFlexGrid.Row, oMSFlexGrid.Col) = pCampoManut.Text
End Sub

Public Sub NavegacaoPosterior(pMSFlexGrid As Object)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    
    Set oMSFlexGrid = pMSFlexGrid
    
    If oMSFlexGrid.Col + 1 < oMSFlexGrid.Cols Then
        GoSub ProximaColuna
    Else
        GoSub ProximaLinha
    End If
    
    Exit Sub
    
ProximaColuna:
    Do
        oMSFlexGrid.Col = oMSFlexGrid.Col + 1
    Loop Until oMSFlexGrid.ColWidth(oMSFlexGrid.Col) > 0 Or (oMSFlexGrid.Col + 1 = oMSFlexGrid.Cols)
    
    If oMSFlexGrid.Col + 1 = oMSFlexGrid.Cols And oMSFlexGrid.ColWidth(oMSFlexGrid.Col) = 0 Then GoSub ProximaLinha
Return

ProximaLinha:
    If oMSFlexGrid.Row + 1 < oMSFlexGrid.Rows Then
        oMSFlexGrid.Row = oMSFlexGrid.Row + 1
        
        oMSFlexGrid.Col = 0
        If oMSFlexGrid.ColWidth(oMSFlexGrid.Col) = 0 Then GoSub ProximaColuna
    End If
Return
End Sub

Public Sub NavegacaoAnterior(pMSFlexGrid As Object)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    
    Set oMSFlexGrid = pMSFlexGrid
    
    If oMSFlexGrid.Col - 1 >= 0 Then
        GoSub ColunaAnterior
    Else
        GoSub LinhaAnterior
    End If
    
    Exit Sub
    
ColunaAnterior:
    Do
        oMSFlexGrid.Col = oMSFlexGrid.Col - 1
    Loop Until oMSFlexGrid.ColWidth(oMSFlexGrid.Col) > 0 Or (oMSFlexGrid.Col - 1 = -1)
    
    If oMSFlexGrid.Col - 1 = -1 And oMSFlexGrid.ColWidth(oMSFlexGrid.Col) = 0 Then GoSub LinhaAnterior
Return

LinhaAnterior:
    If oMSFlexGrid.Row - 1 >= 1 Then
        oMSFlexGrid.Row = oMSFlexGrid.Row - 1
        
        oMSFlexGrid.Col = oMSFlexGrid.Cols - 1
        If oMSFlexGrid.ColWidth(oMSFlexGrid.Col) = 0 Then GoSub ColunaAnterior
    End If
Return
End Sub

Public Sub CorRow(pMSFlexGrid As Object, pLinha As Long, pCor As ADRRBR_SIS_Basico.eCor)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    Dim lColuna     As Integer
    
    Set oMSFlexGrid = pMSFlexGrid
    
    For lColuna = 1 To oMSFlexGrid.Cols - 1
        CorCol oMSFlexGrid, pLinha, lColuna, pCor
    Next lColuna
End Sub

Public Sub CorCol(pMSFlexGrid As Object, pLinha As Long, pColuna As Integer, pCor As ADRRBR_SIS_Basico.eCor)
    Dim oMSFlexGrid As MSFlexGridLib.MSFlexGrid
    
    Set oMSFlexGrid = pMSFlexGrid
    
    oMSFlexGrid.Row = pLinha
    oMSFlexGrid.Col = pColuna
    oMSFlexGrid.CellForeColor = pCor
End Sub

Public Sub LimpaLinha(pMSFlexGrid As Object, pLinha As Long, Optional pColunasFixasDelimVirgula As String)
    Dim oMSFlexGrid     As MSFlexGridLib.MSFlexGrid
    Dim rsAux           As New ADODB.Recordset
    Dim iColuna         As Integer
    Dim sColunasFixas() As String
    
    Set oMSFlexGrid = pMSFlexGrid

    rsAux.Fields.Append "ColunaFixa", adVarChar, "4"
    rsAux.Open

    If InStr(1, pColunasFixasDelimVirgula, ",") > 0 Then
        sColunasFixas() = Split(pColunasFixasDelimVirgula, ",")
        
        For iColuna = 0 To UBound(sColunasFixas())
            rsAux.AddNew
            rsAux!ColunaFixa = Val(sColunasFixas(iColuna))
            rsAux.Update
        Next iColuna
    End If
    
    For iColuna = 0 To oMSFlexGrid.Cols - 1
        rsAux.Filter = ""
        rsAux.Filter = "ColunaFixa = " & CStr(iColuna)
        If rsAux.EOF Then oMSFlexGrid.TextMatrix(pLinha, iColuna) = Empty
    Next iColuna
End Sub

