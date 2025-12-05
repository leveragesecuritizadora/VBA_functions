SELECT
    FORMAT(
      CONVERT(date, LEFT(jo.relatorio_id, 4) + RIGHT(jo.relatorio_id, 2) + '01'),
          'dd/MM/yyyy'
    ) + ' - ' + 
    nome_emissao.nome + ' - ' + CASE nome_unidade.subordinacao  -- Série
      WHEN 'Sênior' THEN 'senior'
      WHEN 'Subordinada' THEN 'subordinada'
      ELSE 'senior'
      END
    AS 'Identificador',
    qntd_cotas.valor_int AS 'Qntd Cotas Dia Anterior ao Pagamento',
    qntd_cotas.valor_int * jo.valor_decimal8 AS 'Juros Ordinário',
    qntd_cotas.valor_int * je.valor_decimal8 AS 'Juros Extraordinário',
    qntd_cotas.valor_int * ao.valor_decimal8 AS 'Amortização Ordinária',
    qntd_cotas.valor_int * ae.valor_decimal8 AS 'Amortização Extraordinário',
    (jo.valor_decimal8 + je.valor_decimal8 + ao.valor_decimal8 + ae.valor_decimal8) * qntd_cotas.valor_int AS 'PMT',
    (jo.valor_decimal8 + je.valor_decimal8) * qntd_cotas.valor_int AS 'PMTsemAMEX'
FROM (
    SELECT -- qntd de cotas
          *
    FROM 
        DW.Fato
    WHERE
        dim_fato_id = 2
        AND dim_indicador_id = 7
) as qntd_cotas
JOIN (
    SELECT 
	    dim_emissao_id,
	    SUBSTRING(emissao_nome, CHARINDEX(' ', emissao_nome) + 1, LEN(emissao_nome)) AS nome
    FROM 
	    DW.DimEmissao
) AS nome_emissao
ON 
    qntd_cotas.dim_emissao_id = nome_emissao.dim_emissao_id
JOIN (
    SELECT -- JO
      *
    FROM DW.Fato
    WHERE
      dim_fato_id = 2
      AND dim_indicador_id = 3
) as jo
ON 
    qntd_cotas.dim_emissao_id = jo.dim_emissao_id
    AND qntd_cotas.relatorio_id = jo.relatorio_id
    AND qntd_cotas.dim_split_id = jo.dim_split_id
JOIN (
    SELECT -- JE
      *
    FROM DW.Fato
    WHERE
      dim_fato_id = 2
      AND dim_indicador_id = 4
) as je
ON 
    jo.dim_emissao_id = je.dim_emissao_id
    AND jo.relatorio_id = je.relatorio_id
    AND jo.dim_split_id = je.dim_split_id
JOIN (
    SELECT -- AO
      *
    FROM DW.Fato
    WHERE
      dim_fato_id = 2
      AND dim_indicador_id = 5
) as ao
ON 
    jo.dim_emissao_id = ao.dim_emissao_id
    AND jo.relatorio_id = ao.relatorio_id
    AND jo.dim_split_id = ao.dim_split_id
JOIN (
    SELECT -- AE
      *
    FROM DW.Fato
    WHERE
      dim_fato_id = 2
      AND dim_indicador_id = 6
) as ae
ON 
    jo.dim_emissao_id = ae.dim_emissao_id
    AND jo.relatorio_id = ae.relatorio_id
    AND jo.dim_split_id = ae.dim_split_id
JOIN (
    SELECT 
        dim_emissao_id,
        dim_split_id,
        subordinacao
    FROM
        DW.DimSplit
) AS nome_unidade
ON 
    nome_unidade.dim_emissao_id = qntd_cotas.dim_emissao_id
    AND nome_unidade.dim_split_id = qntd_cotas.dim_split_id
WHERE 
    qntd_cotas.valor_int IS NOT NULL
    AND jo.valor_decimal8 IS NOT NULL
    AND je.valor_decimal8 IS NOT NULL
    AND ao.valor_decimal8 IS NOT NULL
    AND ae.valor_decimal8 IS NOT NULL
