VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico"
   ClientHeight    =   7515
   ClientLeft      =   3030
   ClientTop       =   2265
   ClientWidth     =   13380
   Icon            =   "frmHistorico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   13380
   Begin VB.Frame fraConsulta 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   7515
      Left            =   0
      TabIndex        =   2
      Tag             =   "NF"
      Top             =   -15
      Width           =   17355
      Begin MSComctlLib.ListView lvwHist 
         Height          =   4875
         Left            =   90
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8599
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   4210752
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin RichTextLib.RichTextBox rtbObservacoes 
         Height          =   1665
         Left            =   135
         TabIndex        =   1
         Top             =   5415
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   2937
         _Version        =   393217
         TextRTF         =   $"frmHistorico.frx":08CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblInforma 
         AutoSize        =   -1  'True
         Caption         =   "Observações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   5835
         TabIndex        =   5
         Top             =   5175
         Width           =   1950
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   270
         Left            =   -180
         TabIndex        =   4
         Tag             =   "LISTA"
         Top             =   0
         Width           =   13590
      End
      Begin VB.Label lblQtdeTrab 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "99 Registro(s) de Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   10815
         TabIndex        =   3
         Tag             =   "LISTA"
         Top             =   7155
         Width           =   2370
      End
   End
End
Attribute VB_Name = "frmHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTituloHist     As String
Private rsRegistrosHist As ADODB.Recordset
 
Const sNomePadrao = "Histórico"
 
Public Property Let TituloHist(ByVal vNewValue As String)
    sTituloHist = vNewValue
End Property
Public Property Get TituloHist() As String
    TituloHist = sTituloHist
End Property
 
Public Property Set RegistrosHist(ByVal vNewValue As ADODB.Recordset)
    Set rsRegistrosHist = vNewValue
End Property
Public Property Get RegistrosHist() As ADODB.Recordset
    Set RegistrosHist = rsRegistrosHist
End Property
 
Private Sub Form_Load()
    If Trim(sTituloHist) <> Empty Then
        lblTitulo.Caption = sTituloHist
    End If
    
    InicializaTela
End Sub

Private Sub Form_Activate()
    If rsRegistrosHist Is Nothing Then
        Retorna
        Exit Sub
    End If

    oBasico.Geral.PosicionaTela Screen, Me

    MontaListaHist
    
    lvwHist.SetFocus
End Sub

Private Sub lvwHist_GotFocus()
    If lvwHist.ListItems.Count = 0 Then Exit Sub
    
    If lvwHist.SelectedItem.Index = 0 Then oBasico.LV.PosicionaRow_Indice lvwHist, 1, True
    lvwHist_ItemClick lvwHist.ListItems.Item(lvwHist.SelectedItem.Index)
End Sub

Private Sub lvwHist_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvwHist.ListItems.Count = 0 Then Exit Sub
    
    rtbObservacoes.Text = lvwHist.ListItems(lvwHist.SelectedItem.Index).SubItems(10)
End Sub

Private Sub lvwHist_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwHist
       .Sorted = True
       If .SortOrder = lvwAscending Then
          .SortOrder = lvwDescending
       Else
          .SortOrder = lvwAscending
       End If
       .SortKey = ColumnHeader.Index - 1
       .Sorted = False
    End With
End Sub

Sub CriaColHist()
    oBasico.LV.AddCol lvwHist, "pk_trabalho_fase_status", "0", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "fk_trabalho_fase_status_tipo", "0", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "Status", "2300", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "pk_trabalho", "0", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "Número Pedido", "1700", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "pk_fase", "0", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "Número Fase", "1200", Centro_LV
    oBasico.LV.AddCol lvwHist, "Fase", "2500", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "fk_operador", "0", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "Colaborador", "2000", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "Obs.Status", "0", Esquerda_LV
    oBasico.LV.AddCol lvwHist, "Dt.Incl.Status", "1700", Centro_LV
    oBasico.LV.AddCol lvwHist, "Interr.LEAD TIME?", "1500", Centro_LV
End Sub

Sub MontaListaHist()
    Screen.MousePointer = vbHourglass
    
    While Not rsRegistrosHist.EOF
        lvwHist.ListItems.Add
        
        With lvwHist.ListItems(lvwHist.ListItems.Count)
            .Text = oBasico.Geral.TrocaNuLL(rsRegistrosHist("pk_trabalho_fase_status"), Empty)
            .SubItems(1) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("fk_trabalho_fase_status_tipo"), Empty)
            .SubItems(2) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Status"), Empty)
            .SubItems(3) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("pk_trabalho"), Empty)
            .SubItems(4) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Número Pedido"), Empty)
            .SubItems(5) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("pk_fase"), Empty)
            .SubItems(6) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Número Fase"), Empty)
            .SubItems(7) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Fase"), Empty)
            .SubItems(8) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("fk_operador"), Empty)
            .SubItems(9) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Colaborador"), Empty)
            .SubItems(10) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Obs.Status"), Empty)
            .SubItems(11) = Format(oBasico.Geral.TrocaNuLL(rsRegistrosHist("Dt.Incl.Status"), Empty), "dd/mm/yyyy hh:mm:ss")
            .SubItems(12) = oBasico.Geral.TrocaNuLL(rsRegistrosHist("Interrompe LEAD TIME?"), Empty)
        End With
        
        rsRegistrosHist.MoveNext
    Wend

    If lvwHist.ListItems.Count = 1 Then
        lblQtdeTrab.Caption = lvwHist.ListItems.Count & " Registro"
    ElseIf lvwHist.ListItems.Count > 1 Then
        lblQtdeTrab.Caption = lvwHist.ListItems.Count & " Registros"
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Retorna ao Formulário de Chamada sem Nenhum Registro Selecionado.
    If KeyCode = vbKeyEscape Then
        Retorna
        Exit Sub
    End If
End Sub

Sub InicializaTela()
    MousePointer = vbHourglass
    
    CriaColHist
    
    rtbObservacoes.BackColor = AzulGelo
    rtbObservacoes.Locked = True
    
    MousePointer = vbDefault
End Sub

Sub Retorna()
    Set rsRegistrosHist = Nothing
    Unload Me
End Sub


