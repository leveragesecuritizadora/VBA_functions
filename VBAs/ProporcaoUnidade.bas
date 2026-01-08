Attribute VB_Name = "ProporcaoUnidade"
Attribute VB_Name = "ProporcaoUnidade"
Public Function PreencherProporcaoUnidade( _
    nome_unidade As String, _
    Optional mes_offset As Variant = -1, _
    Optional coluna_data As Integer = 2 _
) As Variant

    Dim recebimentosTotais As Variant
    Dim recebimentoUnidade As Variant

    recebimentosTotais = PreencherRecebimentosTotaisTU()
    recebimentoUnidade = PreencherRecebimentosTotais(nome_unidade)

    Debug.Print "recebimentosTotais: "; recebimentosTotais
    Debug.Print "recebimentoUnidade: "; recebimentoUnidade

    PreencherProporcaoUnidade = recebimentoUnidade/recebimentosTotais
End Function
