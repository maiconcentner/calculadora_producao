Private Sub CommandButton2_Click()

Unload Me
formMenu.Show

End Sub

Private Sub CommandButton4_Click()

'calculo do tempo de ciclo


Dim tciclo

tciclo = FormatNumber(tProducao / qntProduzidaTC, 2)

'calculo mod

'modTC = FormatNumber(tDisponivelTC / tciclo, 2)

tempoCiclo.Value = tciclo
'qntMOD.Value = modTC

MsgBox "O TEMPO DE CICLO É DE:  " & tciclo & " MINUTO(S) POR CARRO", vbOKOnly, "RESULTADO TEMPO DE CICLO"
End Sub

Private Sub CommandButton9_Click()
'informativo calculo mao de obra direta

    MsgBox "Para cálculo da MOD (Mão de Obra Direta), é utilizado alguns parâmetros, como: tempo disponível de trabalho (518 minutos), índice de absenteísmo, fator de tolerância, demanda e tempo cronometrado para produzir cada unidade.", vbInformation, "Cálculo da MOD"

End Sub