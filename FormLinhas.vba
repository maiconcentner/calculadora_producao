Private Sub CommandButton2_Click()
Unload Me
formMenu.Show
End Sub

Private Sub CommandButton3_Click()

'Calculando o numero de linhas necessária para atender o Takt

'dimensionando variáveis


Dim Resultado
Dim Resultado2

'calculo
On Error Resume Next

Resultado = FormatNumber((txtTakt.Value * txtDemanda.Value) / (txtTempo.Value), 1)
Resultado2 = WorksheetFunction.RoundUp((txtTakt.Value * txtDemanda.Value) / (txtTempo.Value), 0)

txtResultadoLinha.Value = Resultado

MsgBox "A quantidade mínima de linhas para atender a demanda é de: " & Resultado & Chr(13) & Chr(13) _
& "Arredondando para cima: " & Resultado2 & " Linha(s)", vbOKOnly, "Resultado da Operação"


End Sub
