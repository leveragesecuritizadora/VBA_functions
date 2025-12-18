SELECT
    FORMAT(
        CONVERT(
            date,
            LEFT(fd.relatorio_id, 4) + RIGHT(fd.relatorio_id, 2) + '01'
        ),
        'dd/MM/yyyy'
    ) + ' - ' + ne.nome AS Identificador,

    fd.valor_decimal     AS SaldoFD,
    fd_min.valor_decimal AS SaldoMinFD,
    CASE
        WHEN fd.valor_decimal < fd_min.valor_decimal THEN fd_min.valor_decimal - fd.valor_decimal
        ELSE 0
    END AS RecomposicaoFD,
    fr.valor_decimal     AS SaldoFR,
    fr_min.valor_decimal AS SaldoMinFR,
    CASE
        WHEN fr.valor_decimal < fr_min.valor_decimal THEN fr_min.valor_decimal - fr.valor_decimal
        ELSE 0
    END AS RecomposicaoFR

FROM DW.Fato fd
JOIN DW.Fato fd_min
    ON fd.dim_emissao_id = fd_min.dim_emissao_id
   AND fd.relatorio_id   = fd_min.relatorio_id
   AND fd_min.dim_fato_id = 14
   AND fd_min.dim_indicador_id = 8

JOIN DW.Fato fr
    ON fd.dim_emissao_id = fr.dim_emissao_id
   AND fd.relatorio_id   = fr.relatorio_id
   AND fr.dim_fato_id = 6
   AND fr.dim_indicador_id = 5

JOIN DW.Fato fr_min
    ON fd.dim_emissao_id = fr_min.dim_emissao_id
   AND fd.relatorio_id   = fr_min.relatorio_id
   AND fr_min.dim_fato_id = 14
   AND fr_min.dim_indicador_id = 1

JOIN (
    SELECT 
        dim_emissao_id,
        SUBSTRING(
            emissao_nome,
            CHARINDEX(' ', emissao_nome) + 1,
            LEN(emissao_nome)
        ) AS nome
    FROM DW.DimEmissao
) ne
    ON fd.dim_emissao_id = ne.dim_emissao_id

WHERE
    fd.dim_fato_id = 15
    AND fd.dim_indicador_id = 5;
