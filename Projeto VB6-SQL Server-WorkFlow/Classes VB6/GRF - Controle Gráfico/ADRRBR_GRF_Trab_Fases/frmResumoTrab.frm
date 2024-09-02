VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmResumoTrab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo"
   ClientHeight    =   6450
   ClientLeft      =   4680
   ClientTop       =   3435
   ClientWidth     =   9855
   Icon            =   "frmResumoTrab.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9855
   Begin RichTextLib.RichTextBox rtbResumoTrab 
      Height          =   6210
      Left            =   0
      TabIndex        =   1
      Tag             =   "CAMPO;FM1"
      Top             =   270
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   10954
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmResumoTrab.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Resumo do Trabalho"
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
      Left            =   0
      TabIndex        =   0
      Tag             =   "SELECAO"
      Top             =   0
      Width           =   9885
   End
End
Attribute VB_Name = "frmResumoTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sTituloResumo    As String
Private rsRegistroResumo As ADODB.Recordset
 
Const sNomePadrao = "Resumo"
 
Public Property Let TituloResumo(ByVal vNewValue As String)
    sTituloResumo = vNewValue
End Property
Public Property Get TituloResumo() As String
    TituloResumo = sTituloResumo
End Property
 
Public Property Set RegistroResumo(ByVal vNewValue As ADODB.Recordset)
    Set rsRegistroResumo = vNewValue
End Property
Public Property Get RegistroResumo() As ADODB.Recordset
    Set RegistroResumo = rsRegistroResumo
End Property
 
Private Sub Form_Load()
    If Trim(sTituloResumo) <> Empty Then
        lblTitulo.Caption = sTituloResumo
    End If
    
    InicializaTela
End Sub

Private Sub Form_Activate()
    If rsRegistroResumo Is Nothing Then
        Retorna
        Exit Sub
    End If

    oBasico.Geral.PosicionaTela Screen, Me

    MontaResumoTrab
End Sub

Sub MontaResumoTrab()
    Dim sResumo As String
    
    While Not rsRegistroResumo.EOF
        sResumo = Empty
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                              LEAD TIME do Trabalho                            " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Programado...: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!TB_lead_time_programado, "000:00") & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!TB_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!TB_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                           LEAD TIME da fase RECEPÇÃO                          " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM1_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM1_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM1_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf

        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                             LEAD TIME da fase ARTE                            " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM2_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM2_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM2_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                        LEAD TIME da fase REVISÃO DA ARTE                      " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM3_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM3_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM3_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                       LEAD TIME da fase APROVAÇÃO DA ARTE                     " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM4_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM4_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM4_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                           LEAD TIME da fase RETOQUE                           " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM5_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM5_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM5_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                          LEAD TIME da fase PREPARAÇÃO                         " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM6_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM6_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM6_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                     LEAD TIME da fase REVISÃO DA PREPARAÇÃO                   " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM7_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM7_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM7_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                             LEAD TIME da fase PROVA                           " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM8_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM8_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM8_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                       LEAD TIME da fase APROVAÇÃO DA PROVA                    " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM9_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM9_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM9_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                    LEAD TIME da fase REVISÃO DE PROCEDIMENTO                  " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM10_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM10_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM10_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        sResumo = sResumo & "===============================================================================" & vbCrLf
        sResumo = sResumo & "                     LEAD TIME da fase REVISÃO DIGITAL FINAL                   " & vbCrLf
        sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM11_lead_time_consumo, "000:00") & vbCrLf
        sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM11_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM11_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        'sResumo = sResumo & "===============================================================================" & vbCrLf
        'sResumo = sResumo & "                        LEAD TIME da fase PADRÃO DE CORES                      " & vbCrLf
        'sResumo = sResumo & "===============================================================================" & vbCrLf & vbCrLf
        'sResumo = sResumo & " Horas-Minutos Consumido....: " & oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM12_lead_time_consumo, "000:00") & vbCrLf
        'sResumo = sResumo & " Percentual Consumido.......: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM12_lead_time_perc_consumido, "0"), "###,##0.0") & "%" & vbCrLf
        'sResumo = sResumo & " Percentual Estimado........: " & Format(oBasico.Geral.TrocaNuLL(rsRegistroResumo!FM12_lead_time_perc_definido, "0"), "###,##0.0") & "%" & vbCrLf & vbCrLf
        
        rsRegistroResumo.MoveNext
    Wend
    
    rtbResumoTrab.Text = sResumo
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
    
    rtbResumoTrab.Text = Empty
    
    MousePointer = vbDefault
End Sub

Sub Retorna()
    Set rsRegistroResumo = Nothing
    Unload Me
End Sub

