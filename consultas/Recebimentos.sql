SELECT
    FORMAT(
    CONVERT(date, LEFT(fatos.relatorio_id, 4) + RIGHT(fatos.relatorio_id, 2) + '01'),
        'dd/MM/yyyy'
    ) + ' - ' +
    nome_emissao.nome + ' - ' +
    nome_unidade.split_nome AS 'Identificador',
    SUM(
        CASE 
            WHEN fatos.dim_indicador_id BETWEEN 2 AND 11 THEN fatos.valor_decimal
            ELSE 0
        END
    ) AS 'Recebimentos Antecipado',
    SUM(
        CASE 
            WHEN fatos.dim_indicador_id BETWEEN 12 AND 21 THEN fatos.valor_decimal
            ELSE 0
        END
    ) AS 'Recebimentos Atrasado',
    SUM(
        CASE 
            WHEN fatos.dim_indicador_id = 1 THEN fatos.valor_decimal
            ELSE 0
        END
    ) AS 'Recebimentos em dia',
    SUM(
        CASE 
            WHEN fatos.dim_indicador_id BETWEEN 1 AND 21 THEN fatos.valor_decimal
            ELSE 0
        END
    ) AS 'Recebimentos Total'
FROM (
    SELECT 
        *
    FROM
	    DW.Fato 
    WHERE
        dim_fato_id = 9
) as fatos
JOIN (
    SELECT 
	    dim_emissao_id,
	    SUBSTRING(emissao_nome, CHARINDEX(' ', emissao_nome) + 1, LEN(emissao_nome)) AS nome
    FROM 
	    DW.DimEmissao
) AS nome_emissao
ON 
    fatos.dim_emissao_id = nome_emissao.dim_emissao_id
JOIN (
    SELECT 
        dim_emissao_id,
        dim_split_id,
        split_nome
    FROM
        DW.DimSplit
    WHERE
        split_tipo = 'Unidade'
) AS nome_unidade
ON 
    nome_unidade.dim_emissao_id = fatos.dim_emissao_id
    AND nome_unidade.dim_split_id = fatos.dim_split_id
GROUP BY
    FORMAT(
    CONVERT(date, LEFT(fatos.relatorio_id, 4) + RIGHT(fatos.relatorio_id, 2) + '01'),
        'dd/MM/yyyy'
    ) + ' - ' +
    nome_emissao.nome + ' - ' +
    nome_unidade.split_nome