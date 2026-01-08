Attribute VB_Name = "ProporcaoUnidade"

Public Function ProporcaoUnidade( _
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

    ProporcaoUnidade = recebimentoUnidade/recebimentosTotais
End Function
