SELECT
    FORMAT(
      CONVERT(date, LEFT(vu_corrigido.relatorio_id, 4) + RIGHT(vu_corrigido.relatorio_id, 2) + '01'),
          'dd/MM/yyyy'
    ) + ' - ' + 
    nome_emissao.nome + ' - ' + CASE nome_unidade.subordinacao  -- Série
      WHEN 'Sênior' THEN 'senior'
      WHEN 'Subordinada' THEN 'subordinada'
      ELSE 'senior'
      END
    AS 'Identificador', 
    vu_corrigido.valor_decimal8 AS 'Valor Unitário Corrigido'
FROM (
    SELECT -- qntd de cotas
          *
    FROM 
        DW.Fato
    WHERE
        dim_fato_id = 1
        AND dim_indicador_id = 9
) as vu_corrigido
JOIN (
    SELECT 
	    dim_emissao_id,
	    SUBSTRING(emissao_nome, CHARINDEX(' ', emissao_nome) + 1, LEN(emissao_nome)) AS nome
    FROM 
	    DW.DimEmissao
) AS nome_emissao
ON 
    vu_corrigido.dim_emissao_id = nome_emissao.dim_emissao_id
JOIN (
    SELECT 
        dim_emissao_id,
        dim_split_id,
        subordinacao
    FROM
        DW.DimSplit
) AS nome_unidade
ON 
    nome_unidade.dim_emissao_id = vu_corrigido.dim_emissao_id
    AND nome_unidade.dim_split_id = vu_corrigido.dim_split_id
WHERE 
    vu_corrigido.valor_decimal8 IS NOT NULL
