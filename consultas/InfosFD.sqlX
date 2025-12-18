SELECT
    FORMAT(
    CONVERT(date, LEFT(saldo.relatorio_id, 4) + RIGHT(saldo.relatorio_id, 2) + '01'),
        'dd/MM/yyyy'
    ) + ' - ' +
    nome_emissao.nome AS 'Identificador',
    saldo.valor_decimal AS SaldoFD,
	saldo_min.valor_decimal AS SaldoMinFD
FROM (
	SELECT 
		*
	FROM
		DW.Fato
	WHERE
		dim_fato_id = 15
		AND dim_indicador_id = 5
) AS saldo
JOIN (
	SELECT 
		*
	FROM
		DW.Fato
	WHERE
		dim_fato_id = 14
		AND dim_indicador_id = 8
) AS saldo_min
ON 
	saldo.dim_emissao_id = saldo_min.dim_emissao_id
	AND saldo.relatorio_id = saldo_min.relatorio_id
JOIN (
    SELECT 
	    dim_emissao_id,
	    SUBSTRING(emissao_nome, CHARINDEX(' ', emissao_nome) + 1, LEN(emissao_nome)) AS nome
    FROM 
	    DW.DimEmissao
) AS nome_emissao
ON 
    saldo.dim_emissao_id = nome_emissao.dim_emissao_id
