Private Sub CommandButton1_Click()
On Error Resume Next
'calculo do takt time

Dim resultadotakttime

resultadotakttime = FormatNumber((tDisponivel.Value) / (dDiaria / nPosicoes), 2)

resultadoTakt.Value = resultadotakttime

MsgBox "O TAKT TIME PARA ESTE CASO É DE:  " & resultadotakttime & " MINUTOS/HORAS POR CARRO", vbOKOnly, "RESULTADO TAKT"
End Sub

Private Sub CommandButton2_Click()
Unload Me
formMenu.Show
End Sub

Private Sub CommandButton6_Click()

'informativo takt time

MsgBox "Takt Time é o ritmo no qual você precisa completar um produto para suprir a demanda do consumidor. Por exemplo, uma fábrica trabalha 8 horas/dia (480 minutos) e a demanda do mercado é de 120 unidades/dia. Desta forma, o takt time é de 4 minutos (480/120).", vbInformation, "TAKT TIME"

End Sub