SELECT
    FORMAT(
    CONVERT(date, LEFT(antecipacoes.relatorio_id, 4) + RIGHT(antecipacoes.relatorio_id, 2) + '01'),
        'dd/MM/yyyy'
    ) + ' - ' +  CASE
        WHEN antecipacoes.dim_emissao_id = 1 AND antecipacoes.dim_split_id = 6 THEN 'Maranhão - Parque Jardins'
        WHEN antecipacoes.dim_emissao_id = 1 AND antecipacoes.dim_split_id = 7 THEN 'Maranhão - Cidade Nova'

        WHEN antecipacoes.dim_emissao_id = 3 AND antecipacoes.dim_split_id = 7 THEN 'R&O - Villa Flora'
        WHEN antecipacoes.dim_emissao_id = 3 AND antecipacoes.dim_split_id = 8 THEN 'R&O - Monte Belo'

        WHEN antecipacoes.dim_emissao_id = 4 AND antecipacoes.dim_split_id = 3 THEN 'Smart Sabiás - Euroville Mall'
        WHEN antecipacoes.dim_emissao_id = 4 AND antecipacoes.dim_split_id = 4 THEN 'Smart Sabiás - Parque Residencial Sabiás'
        WHEN antecipacoes.dim_emissao_id = 4 AND antecipacoes.dim_split_id = 5 THEN 'Smart Sabiás - Smart City Indaiá'
        WHEN antecipacoes.dim_emissao_id = 4 AND antecipacoes.dim_split_id = 6 THEN 'Smart Sabiás - Empreendimentos'

        WHEN antecipacoes.dim_emissao_id = 5 AND antecipacoes.dim_split_id = 3 THEN 'Impegno - Provincia Di Salerno'

        WHEN antecipacoes.dim_emissao_id = 6 AND antecipacoes.dim_split_id = 5 THEN 'Fazenda Ranchão - Fazenda Ranchão'

        WHEN antecipacoes.dim_emissao_id = 7 AND antecipacoes.dim_split_id = 7 THEN 'Garden Ville - Garden Ville'
        WHEN antecipacoes.dim_emissao_id = 7 AND antecipacoes.dim_split_id = 8 THEN 'Garden Ville - Quintas'

        WHEN antecipacoes.dim_emissao_id = 8 AND antecipacoes.dim_split_id = 3 THEN 'Porto Real - Loteamento PR'
        WHEN antecipacoes.dim_emissao_id = 8 AND antecipacoes.dim_split_id = 4 THEN 'Porto Real - Loteamento PR2'

        WHEN antecipacoes.dim_emissao_id = 9 AND antecipacoes.dim_split_id = 3 THEN 'Viva Urban - Urban Caribe Design'
        WHEN antecipacoes.dim_emissao_id = 9 AND antecipacoes.dim_split_id = 4 THEN 'Viva Urban - Veleiro Urban Design'

        ELSE NULL
    END AS 'Identificador',
    total.valor as 'Recebimento no mês',
    antecipacoes.valor as 'Recebimento Antecipado',
    atrasos.valor as 'Recebimento Atrasado',
    total.valor+antecipacoes.valor+atrasos.valor AS 'Recebimentos Total'
FROM (
    SELECT 
        dim_emissao_id,
        dim_split_id, 
        relatorio_id, 
        SUM(valor_decimal) AS 'valor'
    FROM DW.Fato
    WHERE 
        dim_indicador_id BETWEEN 2 AND 11
        AND dim_fato_id = 9
        AND dim_split_id BETWEEN 6 AND 7
    GROUP BY
        dim_emissao_id,
        dim_split_id,
        relatorio_id
) AS antecipacoes
JOIN (
    SELECT 
        dim_emissao_id,
        dim_split_id, 
        relatorio_id, 
        SUM(valor_decimal) AS 'valor'
    FROM DW.Fato
    WHERE 
        dim_indicador_id BETWEEN 12 AND 21
        AND dim_fato_id = 9
        AND dim_split_id BETWEEN 6 AND 7
    GROUP BY
        dim_emissao_id,
        dim_split_id,
        relatorio_id
) AS atrasos
ON 
    antecipacoes.dim_split_id = atrasos.dim_split_id
    AND antecipacoes.relatorio_id = atrasos.relatorio_id
JOIN (
    SELECT 
	    dim_split_sk_id,
	    relatorio_id,
	    SUM(valor_decimal) AS 'valor'
    FROM 
	    DW.Fato
    WHERE
	    dim_fato_id = 9
	    AND dim_indicador_id = 1
    GROUP BY
	    dim_split_sk_id,
	    relatorio_id
) AS total
ON 
    antecipacoes.dim_split_id = total.dim_split_sk_id
    AND antecipacoes.relatorio_id = total.relatorio_id
ORDER BY 
    antecipacoes.dim_split_id,
    antecipacoes.relatorio_id