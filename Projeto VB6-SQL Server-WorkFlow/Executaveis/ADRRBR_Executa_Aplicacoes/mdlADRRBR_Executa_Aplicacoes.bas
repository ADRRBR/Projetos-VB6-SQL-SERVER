Attribute VB_Name = "mdlADRRBR_Executa_Aplicacoes"
Option Explicit

Dim oExecAplicacao As Object

Sub Main()
    If App.PrevInstance Then
        MsgBox "Este programa j� est� em execu��o!", vbExclamation, "Aten��o"
        End
        Exit Sub
    End If
    
    Set oExecAplicacao = CreateObject("ADRRBR_APL_Aplicacoes.clsAPL_Aplicacoes")
    
    oExecAplicacao.TipoBancoDados = SQL_Server
    oExecAplicacao.Carrega
    
    If oExecAplicacao.Status = Erro Then
        MsgBox oExecAplicacao.Mensagem, vbCritical, "Aten��o"
        Exit Sub
    End If
End Sub
