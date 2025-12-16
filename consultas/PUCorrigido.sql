SELECT
    FORMAT(
      CONVERT(date, LEFT(fat_indice.relatorio_id, 4) + RIGHT(fat_indice.relatorio_id, 2) + '01'),
          'dd/MM/yyyy'
    ) + ' - ' + 
    nome_emissao.nome + ' - ' + CASE subord.subordinacao  -- Série
      WHEN 'Sênior' THEN 'senior'
      WHEN 'Subordinada' THEN 'subordinada'
      ELSE 'senior'
      END
    AS 'Identificador', 
    val_base.valor_decimal8 AS 'Valor Base',
    fat_indice.valor_decimal8 AS 'Fator índice',
    fat_indice.valor_decimal8 * val_base.valor_decimal8 AS 'Valor Unitário Corrigido'
FROM (
    SELECT -- fator indice
          *
    FROM 
        DW.Fato
    WHERE
        dim_fato_id = 2
        AND dim_indicador_id = 11
) AS fat_indice
JOIN (
    SELECT -- valor base
          *
    FROM 
        DW.Fato
    WHERE
        dim_fato_id = 2
        AND dim_indicador_id = 8
) AS val_base
ON
    fat_indice.dim_emissao_id = val_base.dim_emissao_id
    AND fat_indice.dim_split_id = fat_indice.dim_split_id
    AND fat_indice.relatorio_id = val_base.relatorio_id
JOIN (
    SELECT 
	    dim_emissao_id,
	    SUBSTRING(emissao_nome, CHARINDEX(' ', emissao_nome) + 1, LEN(emissao_nome)) AS nome
    FROM 
	    DW.DimEmissao
) AS nome_emissao
ON 
    fat_indice.dim_emissao_id = nome_emissao.dim_emissao_id
JOIN (
    SELECT 
        dim_emissao_id,
        dim_split_id,
        subordinacao
    FROM
        DW.DimSplit
) AS subord
ON 
    subord.dim_emissao_id = fat_indice.dim_emissao_id
    AND subord.dim_split_id = fat_indice.dim_split_id
WHERE 
    fat_indice.valor_decimal8 IS NOT NULL
