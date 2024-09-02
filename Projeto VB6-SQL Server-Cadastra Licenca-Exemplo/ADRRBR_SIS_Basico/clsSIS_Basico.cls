VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSIS_Basico"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oGeral     As ADRRBR_SIS_Basico.clsSIS_Geral
Private oBD        As ADRRBR_SIS_Basico.clsSIS_BD
Private oFG        As ADRRBR_SIS_Basico.clsSIS_FG
Private oLV        As ADRRBR_SIS_Basico.clsSIS_LV
Private oTV        As ADRRBR_SIS_Basico.clsSIS_TV
Private oDataHora  As ADRRBR_SIS_Basico.clsSIS_DataHora
Private oValor     As ADRRBR_SIS_Basico.clsSIS_Valor
Private oCNPJCPF   As ADRRBR_SIS_Basico.clsSIS_CNPJCPF
Private oGravaLog  As ADRRBR_SIS_Basico.clsSIS_GravaLog
Private oMDB       As ADRRBR_SIS_Basico.clsSIS_MDB
Private oExcel     As ADRRBR_SIS_Basico.clsSIS_Excel
Private oWord      As ADRRBR_SIS_Basico.clsSIS_Word
Private oDocClasse As ADRRBR_SIS_Basico.clsSIS_DocClasse

Private Const sClasse = "SIS_Basico"

Private Sub Class_Initialize()
    Set oGeral = New ADRRBR_SIS_Basico.clsSIS_Geral
    Set oBD = New ADRRBR_SIS_Basico.clsSIS_BD
    Set oFG = New ADRRBR_SIS_Basico.clsSIS_FG
    Set oLV = New ADRRBR_SIS_Basico.clsSIS_LV
    Set oTV = New ADRRBR_SIS_Basico.clsSIS_TV
    Set oDataHora = New ADRRBR_SIS_Basico.clsSIS_DataHora
    Set oValor = New ADRRBR_SIS_Basico.clsSIS_Valor
    Set oCNPJCPF = New ADRRBR_SIS_Basico.clsSIS_CNPJCPF
    Set oGravaLog = New ADRRBR_SIS_Basico.clsSIS_GravaLog
    Set oMDB = New ADRRBR_SIS_Basico.clsSIS_MDB
    Set oExcel = New ADRRBR_SIS_Basico.clsSIS_Excel
    Set oWord = New ADRRBR_SIS_Basico.clsSIS_Word
    Set oDocClasse = New ADRRBR_SIS_Basico.clsSIS_DocClasse
End Sub

Private Sub Class_Terminate()
    Set oGeral = Nothing
    Set oBD = Nothing
    Set oFG = Nothing
    Set oLV = Nothing
    Set oTV = Nothing
    Set oDataHora = Nothing
    Set oValor = Nothing
    Set oCNPJCPF = Nothing
    Set oGravaLog = Nothing
    Set oMDB = Nothing
    Set oExcel = Nothing
    Set oWord = Nothing
    Set oDocClasse = Nothing

    Set oBasico = Nothing
End Sub

Public Property Get Geral() As ADRRBR_SIS_Basico.clsSIS_Geral
    Set Geral = oGeral
End Property

Public Property Get BD() As ADRRBR_SIS_Basico.clsSIS_BD
    Set BD = oBD
End Property

Public Property Get FG() As ADRRBR_SIS_Basico.clsSIS_FG
    Set FG = oFG
End Property

Public Property Get LV() As ADRRBR_SIS_Basico.clsSIS_LV
    Set LV = oLV
End Property

Public Property Get TV() As ADRRBR_SIS_Basico.clsSIS_TV
    Set TV = oTV
End Property

Public Property Get DataHora() As ADRRBR_SIS_Basico.clsSIS_DataHora
    Set DataHora = oDataHora
End Property

Public Property Get Valor() As ADRRBR_SIS_Basico.clsSIS_Valor
    Set Valor = oValor
End Property

Public Property Get CNPJCPF() As ADRRBR_SIS_Basico.clsSIS_CNPJCPF
    Set CNPJCPF = oCNPJCPF
End Property

Public Property Get Log() As ADRRBR_SIS_Basico.clsSIS_GravaLog
    Set Log = oGravaLog
End Property

Public Property Get MDB() As ADRRBR_SIS_Basico.clsSIS_MDB
    Set MDB = oMDB
End Property

Public Property Get Excel() As ADRRBR_SIS_Basico.clsSIS_Excel
    Set Excel = oExcel
End Property

Public Property Get Word() As ADRRBR_SIS_Basico.clsSIS_Word
    Set Word = oWord
End Property

Public Property Get DocClasse() As ADRRBR_SIS_Basico.clsSIS_DocClasse
    Set DocClasse = oDocClasse
End Property

