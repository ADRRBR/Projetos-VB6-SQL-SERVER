VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Valor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub FormatoValor(pCampoValor As Object, pCasasDecimais As Integer)
    Dim oCampoValor      As TextBox
    Dim sValor           As String
    Dim iPartes          As Integer
    Dim sValorInteiro()  As String
    Dim sValorInteiroAux As String
    Dim sMilhares()      As String
    Dim sMilharesAux     As String
    
    Set oCampoValor = pCampoValor
    
    sValor = oCampoValor.Text
    sValor = Replace(sValor, ".", "")
    sValor = Replace(sValor, ",", "")
    
    'Verifica Decimais
    If Len(sValor) > pCasasDecimais Then 'Insere Vírgula
        sValor = Left(sValor, Len(sValor) - pCasasDecimais) & "," & Right(sValor, pCasasDecimais)
    End If
    
    'Verifica Milhares
    sValorInteiro = Split(sValor, ",")
    If UBound(sValorInteiro()) > 0 Then
        sValorInteiroAux = sValorInteiro(0)
        ReDim Preserve sMilhares(0)
        Do
            If Len(sValorInteiroAux) > 3 Then
                sMilharesAux = Right(sValorInteiroAux, 3)
                sMilhares(UBound(sMilhares())) = sMilharesAux
                
                sValorInteiroAux = Left(sValorInteiroAux, Len(sValorInteiroAux) - 3)
            Else
                sMilhares(UBound(sMilhares())) = sValorInteiroAux
                sValorInteiroAux = Empty
            End If
            
            If Len(sValorInteiroAux) > 0 Then
                ReDim Preserve sMilhares(UBound(sMilhares()) + 1)
            End If
        Loop Until Len(sValorInteiroAux) = 0
        
        'Compõe Milhares (Separando por Ponto)
        sValor = Empty
        For iPartes = UBound(sMilhares()) To 0 Step -1
            sValor = sValor & sMilhares(iPartes) & "."
        Next iPartes
        
        'Compõe Milhares com Decimais (Separando Por Vírgula)
        sValor = Left(sValor, Len(sValor) - 1) & "," & sValorInteiro(1)
    End If
    
    'Atualiza o Campo Valor
    If sValor = Empty Then
        SelecionaValor oCampoValor, pCasasDecimais
    Else
        oCampoValor = sValor
        oCampoValor.SelStart = Len(oCampoValor)
    End If
End Sub

Public Sub SelecionaValor(pCampoValor As Object, pCasasDecimais As Integer)
    Dim oCampoValor     As TextBox
    Dim iPosicaoValida  As Integer
    Dim iPosicao        As Integer
    Dim sCaracter       As String * 1
    
    Set oCampoValor = pCampoValor
    
    'Consiste Tamanho Máximo do Campo de Valor (Propriedade MaxLenght)
    If oCampoValor.MaxLength < pCasasDecimais + 2 Then
        oCampoValor.MaxLength = pCasasDecimais + 2
    End If
    
    'Retira Zeros à Esquerda
    If Left(oCampoValor, 1) = "0" Then
        iPosicaoValida = InStr(1, oCampoValor, ",")
        If iPosicaoValida = 0 Then
            iPosicaoValida = Len(oCampoValor)
        Else
            iPosicaoValida = iPosicaoValida - 1
        End If
        For iPosicao = 1 To iPosicaoValida
            sCaracter = Mid(oCampoValor, iPosicao, 1)
            If sCaracter <> "0" And sCaracter <> "." Then
                iPosicaoValida = iPosicao
                Exit For
            End If
        Next iPosicao
        oCampoValor = Mid(oCampoValor, iPosicaoValida, Len(oCampoValor) - iPosicaoValida + 1)
    End If
    
    If Trim(oCampoValor) = Empty Then
        oCampoValor = "0," & Format(0, String(pCasasDecimais, "0"))
    
    ElseIf CCur(oCampoValor) = 0 Then
        oCampoValor = "0," & Format(0, String(pCasasDecimais, "0"))
    
    ElseIf InStr(1, oCampoValor, ",") = 0 Then
        oCampoValor = oCampoValor & "," & String(pCasasDecimais, "0")
    End If
    
    oCampoValor.SelStart = 0
    oCampoValor.SelLength = Len(oCampoValor.Text)
End Sub

Public Function FormataValorGravar(pValorFormatar As Variant) As String
    Dim sDecimais As String
    Dim iAux      As Integer
    Dim sAux()    As String
    Dim sAux2     As String
       
    sAux2 = Replace(pValorFormatar, ",", "")
    sAux2 = Replace(sAux2, ".", "")
   
    If Not IsNumeric(sAux2) Then
        FormataValorGravar = pValorFormatar
        Exit Function
    End If
       
    pValorFormatar = Replace(pValorFormatar, ",", ".")
   
    sAux() = Split(pValorFormatar, ".")
   
    If UBound(sAux()) > 0 Then
        sDecimais = sAux(UBound(sAux()))
       
        For iAux = LBound(sAux()) To UBound(sAux()) - 1
            FormataValorGravar = FormataValorGravar & sAux(iAux)
        Next iAux
       
        FormataValorGravar = FormataValorGravar & "." & sDecimais
    Else
        FormataValorGravar = sAux(0) & ".00"
    End If
End Function
