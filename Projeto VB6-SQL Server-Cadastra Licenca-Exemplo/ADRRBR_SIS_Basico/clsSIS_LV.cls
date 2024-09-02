VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_LV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum eAlinhamentoLV
    Esquerda_LV = lvwColumnLeft
    Direita_LV = lvwColumnRight
    Centro_LV = lvwColumnCenter
End Enum

Public Sub AddCol(pListView As Variant, pColuna As Variant, pTamanho As Variant, Optional pAlinhamentoLV As ADRRBR_SIS_Basico.eAlinhamentoLV)
    Dim oListView As MSComctlLib.ListView
    
    Set oListView = pListView
    
    With oListView
        .ColumnHeaders.Add , , pColuna
        .ColumnHeaders(.ColumnHeaders.Count).Width = pTamanho
        .ColumnHeaders(.ColumnHeaders.Count).Alignment = pAlinhamentoLV
    End With
End Sub

Public Sub AddReg(pListView As Variant, pRSListView As ADODB.Recordset, pInicializarListView As Boolean)
    Dim oListView As MSComctlLib.ListView
    Dim lCol      As Long
    Dim sConteudo As String
    
    Set oListView = pListView
    
    If pRSListView.RecordCount = 0 Then Exit Sub
    
    If pInicializarListView Then
        oListView.ListItems.Clear
        oListView.ColumnHeaders.Clear
    End If
    
    If oListView.ColumnHeaders.Count = 0 Then
        For lCol = 0 To pRSListView.Fields.Count - 1
            Select Case lCol
                Case 0
                    oBasico.LV.AddCol oListView, pRSListView.Fields(lCol).Name, 1000
                
                Case Else
                    oBasico.LV.AddCol oListView, pRSListView.Fields(lCol).Name, 1000, Esquerda_LV
            End Select
        Next lCol
    End If
    
    pRSListView.MoveFirst
    
    While Not pRSListView.EOF
        oListView.ListItems.Add
        
        For lCol = 0 To pRSListView.Fields.Count - 1
            If pRSListView.Fields(lCol).Type = adCurrency Then
                sConteudo = Format(oBasico.Geral.TrocaNuLL(pRSListView.Fields(lCol).Value, "0"), "###,###,###,##0.00")
            Else
                sConteudo = oBasico.Geral.TrocaNuLL(pRSListView.Fields(lCol).Value, Empty)
            End If
            
            Select Case lCol
                Case 0
                    oListView.ListItems(oListView.ListItems.Count).Text = sConteudo
                
                Case Else
                    oListView.ListItems(oListView.ListItems.Count).SubItems(lCol) = sConteudo
            End Select
        Next lCol
        
        pRSListView.MoveNext
    Wend
End Sub

Public Sub CorRow(pListView As Variant, pLinha As Variant, pCor As ADRRBR_SIS_Basico.eCor)
    Dim oListView As MSComctlLib.ListView
    Dim lColuna   As Long
    
    Set oListView = pListView
    
    For lColuna = 1 To oListView.ColumnHeaders.Count
        CorCol oListView, pLinha, lColuna, pCor
    Next lColuna
End Sub

Public Sub CorCol(pListView As Variant, pLinha As Variant, pColuna As Variant, pCor As ADRRBR_SIS_Basico.eCor)
    Dim oListView As MSComctlLib.ListView
    
    Set oListView = pListView
    
    With oListView.ListItems(pLinha)
        If pColuna = 1 Then
            .ForeColor = pCor
        Else
            .ListSubItems(pColuna - 1).ForeColor = pCor
        End If
    End With
End Sub

Public Sub Pesquisa_Conteudo(pListView As Variant, pTituloPesq As String)
    Dim oListView As MSComctlLib.ListView
    
    Set oListView = pListView

    Set frmPesqLV.pListView = oListView
    frmPesqLV.pPesquisando = True
    frmPesqLV.pTituloPesq = pTituloPesq
    frmPesqLV.Show vbModal
    
    If Trim(frmPesqLV.pConteudoPesq) <> Empty Then
        With oListView
           .Sorted = True
           .SortOrder = lvwAscending
           .SortKey = frmPesqLV.pColunaPesq - 1
           .Sorted = False
        End With
        
        PosicionaRow_Indice oListView, 1, True
        PosicionaRow_Conteudo oListView, frmPesqLV.pColunaPesq, frmPesqLV.pConteudoPesq, True, frmPesqLV.pPesqContido
    End If
    frmPesqLV.pPesquisando = False
End Sub

Public Function PosicionaRow_Indice(pListView As Variant, pIndicePosicionar As Variant, pPosicionarComFoco As Boolean) As Boolean
    Dim oListView As MSComctlLib.ListView
    
    Set oListView = pListView
    
    PosicionaRow_Indice = False
    
    If oListView.ListItems.Count > 0 Then
        PosicionaRow_Indice = True
        oListView.ListItems(pIndicePosicionar).Selected = True
        
        If pPosicionarComFoco Then
            oListView.SelectedItem.EnsureVisible
            oListView.SetFocus
        End If
    End If
End Function

Public Function PosicionaRow_Conteudo(pListView As Variant, pColuna As Variant, pConteudo As Variant, pPosicionarComFoco As Boolean, Optional pConsiderarConteudoContido As Boolean) As Boolean
    Dim oListView As MSComctlLib.ListView
    Dim lLinha    As Long
    Dim sConteudo As String
    
    Set oListView = pListView
    
    PosicionaRow_Conteudo = False
    
    pConteudo = CStr(Trim(UCase(pConteudo)))
    
    For lLinha = 1 To oListView.ListItems.Count
        If pColuna = 1 Then
            sConteudo = UCase(oListView.ListItems(lLinha).Text)
        Else
            sConteudo = UCase(oListView.ListItems(lLinha).SubItems(pColuna - 1))
        End If
        
        If pConsiderarConteudoContido Then
            If InStr(sConteudo, pConteudo) Then PosicionaRow_Conteudo = True
        Else
            If sConteudo = pConteudo Then PosicionaRow_Conteudo = True
        End If
        
        If PosicionaRow_Conteudo Then
            oListView.ListItems(lLinha).Selected = True
        
            If pPosicionarComFoco Then
                oListView.SelectedItem.EnsureVisible
                oListView.SetFocus
            End If
            Exit For
        End If
    Next lLinha
End Function

Public Sub OrdenaCol(pListView As Variant, pColuna As Long)
    Dim oListView As MSComctlLib.ListView
    
    Set oListView = pListView

    If frmPesqLV.pPesquisando Then Exit Sub
    
    With oListView
       .Sorted = True
       If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
       Else
          .SortOrder = lvwAscending
       End If
       .SortKey = pColuna - 1
       .Sorted = False
    End With
End Sub
