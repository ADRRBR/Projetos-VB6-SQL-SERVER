VERSION 5.00
Begin VB.Form frmPesqLV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa no ListView"
   ClientHeight    =   990
   ClientLeft      =   4410
   ClientTop       =   4935
   ClientWidth     =   3885
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3885
   Begin VB.OptionButton optTipoPesq 
      Caption         =   "Pesquisar texto &Igual"
      Height          =   225
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1965
   End
   Begin VB.OptionButton optTipoPesq 
      Caption         =   "Pesquisar texto &Contido"
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Value           =   -1  'True
      Width           =   1965
   End
   Begin VB.TextBox txtPesq 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   630
      Width           =   3885
   End
   Begin VB.ComboBox cmbColuna 
      BackColor       =   &H00E6E9E2&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "frmPesqLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pListView      As MSComctlLib.ListView 'ListView a Ser Pesquisado
Public pColunaPesq    As Integer              'Número da Coluna de Retorno Para Pesquisa no ListView
Public pConteudoPesq  As String               'Conteudo de Retorno Para Pesquisa no ListView
Public pTituloPesq    As String               'Título da Lista Para a Tela de Pesquisa
Public pPesqContido   As Boolean              'Tipo de Pesquisa (Conteúdo Igual ou Conteúdo Contido)
Public pPesquisando   As Boolean              'Indica que a Pesquisa Está Sendo Executada (Tela de Pesquisa Ativa)

Dim iColuna           As Integer
Dim lIndice           As Long

Private Sub Form_Activate()
    If Not PreparaPesquisa Then
        Unload Me
        Exit Sub
    End If
    
    Me.Caption = "Pesquisa"
    If Trim(pTituloPesq) <> Empty Then Me.Caption = Me.Caption & " - " & pTituloPesq
End Sub
Private Sub cmbColuna_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub txtPesq_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        pColunaPesq = cmbColuna.ItemData(cmbColuna.ListIndex)
        pConteudoPesq = Trim(txtPesq.Text)
        pPesqContido = optTipoPesq(0).Value
        Unload Me
        Exit Sub
    End If
End Sub
Private Sub cmbColuna_Click()
    txtPesq.SetFocus
End Sub
Private Sub optTipoPesq_Click(Index As Integer)
    txtPesq.SetFocus
End Sub
Private Sub optTipoPesq_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        pColunaPesq = Empty
        pConteudoPesq = Empty
        Unload Me
        Exit Sub
    End If
End Sub
Function PreparaPesquisa() As Boolean
    Dim sNomeColuna As String
    
    PreparaPesquisa = True
    cmbColuna.Clear
    
    For iColuna = 1 To pListView.ColumnHeaders.Count
        If pListView.ColumnHeaders(iColuna).Width > 0 Then
            sNomeColuna = pListView.ColumnHeaders(iColuna).Text
            cmbColuna.AddItem sNomeColuna
            cmbColuna.ItemData(cmbColuna.NewIndex) = iColuna
        End If
    Next iColuna
    
    If cmbColuna.ListCount = 0 Then
        PreparaPesquisa = False
        Exit Function
    End If
    
    cmbColuna.ListIndex = 0
End Function

